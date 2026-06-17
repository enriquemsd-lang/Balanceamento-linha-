"""
Simulador de Capacidade e Programação de Produção
Fábrica com máquinas de corte a laser e prensa Hotstamp
"""

import io
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

# ─────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Simulador de Produção",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# INJEÇÃO DE CSS – identidade visual industrial
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* fundo geral */
.stApp { background-color: #0d1117; color: #e6edf3; }

/* sidebar */
[data-testid="stSidebar"] { background-color: #161b22; border-right: 1px solid #30363d; }
[data-testid="stSidebar"] * { color: #c9d1d9 !important; }

/* cabeçalhos */
h1 { font-size: 1.7rem !important; font-weight: 700; color: #58a6ff !important; letter-spacing: -0.5px; }
h2 { font-size: 1.25rem !important; font-weight: 600; color: #79c0ff !important; border-bottom: 1px solid #30363d; padding-bottom: 6px; }
h3 { font-size: 1rem !important; font-weight: 600; color: #d2a8ff !important; }

/* métricas */
[data-testid="metric-container"] {
    background: #161b22;
    border: 1px solid #30363d;
    border-radius: 8px;
    padding: 14px 18px;
}
[data-testid="metric-container"] label { color: #8b949e !important; font-size: 0.75rem !important; text-transform: uppercase; letter-spacing: 0.05em; }
[data-testid="metric-container"] [data-testid="stMetricValue"] { font-family: 'JetBrains Mono', monospace; font-size: 1.6rem !important; color: #f0f6fc !important; }
[data-testid="stMetricDelta"] { font-size: 0.8rem !important; }

/* tabelas */
[data-testid="stDataFrame"] { border: 1px solid #30363d; border-radius: 8px; overflow: hidden; }

/* botão primário */
.stButton > button {
    background: #238636; color: #f0f6fc;
    border: 1px solid #2ea043; border-radius: 6px;
    font-weight: 600; padding: 8px 20px;
    transition: background 0.2s;
}
.stButton > button:hover { background: #2ea043; }

/* download button */
.stDownloadButton > button {
    background: #1f6feb; color: #f0f6fc;
    border: 1px solid #388bfd; border-radius: 6px;
    font-weight: 600; padding: 8px 20px;
}
.stDownloadButton > button:hover { background: #388bfd; }

/* alertas */
.stSuccess { background: #0f2a1a; border: 1px solid #2ea043; border-radius: 6px; }
.stError   { background: #2a0e1a; border: 1px solid #f85149; border-radius: 6px; }
.stWarning { background: #2a1e00; border: 1px solid #d29922; border-radius: 6px; }
.stInfo    { background: #0c1e30; border: 1px solid #388bfd; border-radius: 6px; }

/* chips de status */
.chip-ok         { background:#0f2a1a; color:#3fb950; border:1px solid #2ea043; border-radius:12px; padding:2px 10px; font-size:0.78rem; font-weight:600; }
.chip-sobrecarga { background:#2a0e1a; color:#f85149; border:1px solid #f85149; border-radius:12px; padding:2px 10px; font-size:0.78rem; font-weight:600; }

/* divisor */
hr { border-color: #30363d !important; }

/* upload area */
[data-testid="stFileUploader"] { background: #161b22; border: 1px dashed #30363d; border-radius: 8px; padding: 12px; }

/* expander */
[data-testid="stExpander"] { background:#161b22; border:1px solid #30363d; border-radius:8px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTES – abas e colunas obrigatórias
# ─────────────────────────────────────────────
ABAS_OBRIGATORIAS = {
    "Recursos": [
        "Recurso", "Tipo de recurso", "Centro de trabalho",
        "Turno", "Horas disponíveis por dia", "Eficiência padrão",
        "Setup médio", "Observações",
    ],
    "Produtos": [
        "Código do produto", "Descrição", "Cliente",
        "Tipo de demanda", "Prioridade", "Família",
        "Lote mínimo", "Estoque de segurança", "Prazo padrão",
    ],
    "Roteiros": [
        "Produto", "Recurso permitido", "Tempo de ciclo (s)",
        "Setup (min)", "Rendimento", "Observação técnica",
    ],
    "Demandas": [
        "Ordem", "Produto", "Cliente", "Quantidade",
        "Data de necessidade", "Prioridade", "Destino", "Status",
    ],
    "Calendario": [
        "Recurso", "Data", "Turno", "Minutos disponíveis",
        "Paradas planejadas", "Minutos perdidos", "Capacidade líquida",
    ],
    "Programacao": [
        "Recurso", "Ordem", "Produto", "Quantidade",
        "Início previsto", "Fim previsto", "Sequência",
        "Setup", "Status", "Carga (min)",
    ],
    "Dashboard": ["Carga total", "Capacidade total", "Utilização"],
}

# ─────────────────────────────────────────────
# PALETA PLOTLY (tema escuro industrial)
# ─────────────────────────────────────────────
PLOTLY_TEMPLATE = "plotly_dark"
COR_OK          = "#3fb950"
COR_SOBRECARGA  = "#f85149"
COR_CAPACIDADE  = "#388bfd"
COR_CARGA       = "#d29922"
COR_GARGALO     = "#ff7b72"


# ═══════════════════════════════════════════════
# 1. VALIDAÇÃO DO EXCEL
# ═══════════════════════════════════════════════
def validar_excel(xls: pd.ExcelFile) -> list[str]:
    """
    Verifica se todas as abas e colunas obrigatórias estão presentes.
    Retorna lista de erros (vazia = arquivo válido).
    """
    erros = []
    abas_presentes = set(xls.sheet_names)

    for aba, colunas in ABAS_OBRIGATORIAS.items():
        if aba not in abas_presentes:
            erros.append(f"❌ Aba ausente: **{aba}**")
            continue
        df = xls.parse(aba, nrows=0)   # apenas cabeçalho
        cols_presentes = set(df.columns)
        for col in colunas:
            if col not in cols_presentes:
                erros.append(f"❌ Aba **{aba}** – coluna ausente: `{col}`")
    return erros


# ═══════════════════════════════════════════════
# 2. CARREGAMENTO DAS ABAS
# ═══════════════════════════════════════════════
def carregar_abas(xls: pd.ExcelFile) -> dict[str, pd.DataFrame]:
    """Lê todas as abas em um dicionário de DataFrames."""
    return {aba: xls.parse(aba) for aba in ABAS_OBRIGATORIAS}


# ═══════════════════════════════════════════════
# 3. ENRIQUECIMENTO: DEMANDAS × ROTEIROS
# ═══════════════════════════════════════════════
def enriquecer_demandas(demandas: pd.DataFrame, roteiros: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Cruza Demandas com Roteiros.
    Para produtos com múltiplos recursos permitidos, escolhe
    a rota com menor Tempo de ciclo (s).

    Retorna:
        df_enriquecido : ordens com rota válida
        df_sem_rota    : ordens sem rota encontrada
    """
    # Melhor rota por produto: menor tempo de ciclo
    melhor_rota = (
        roteiros
        .sort_values("Tempo de ciclo (s)")
        .drop_duplicates(subset=["Produto"], keep="first")
        [["Produto", "Recurso permitido", "Tempo de ciclo (s)", "Setup (min)", "Rendimento"]]
    )

    merged = demandas.merge(
        melhor_rota,
        left_on="Produto",
        right_on="Produto",
        how="left",
    )

    # Cálculo da carga
    merged["Carga produção (min)"] = (
        merged["Quantidade"] * merged["Tempo de ciclo (s)"] / 60
        / merged["Rendimento"].replace(0, 1)  # evita divisão por zero
    )
    merged["Carga total (min)"] = merged["Carga produção (min)"] + merged["Setup (min)"].fillna(0)

    sem_rota     = merged[merged["Recurso permitido"].isna()].copy()
    enriquecido  = merged[merged["Recurso permitido"].notna()].copy()

    return enriquecido, sem_rota


# ═══════════════════════════════════════════════
# 4. CARGA POR RECURSO
# ═══════════════════════════════════════════════
def calcular_carga_por_recurso(enriquecido: pd.DataFrame) -> pd.DataFrame:
    """Agrega carga total por recurso."""
    return (
        enriquecido
        .groupby("Recurso permitido", as_index=False)
        .agg(Ordens=("Ordem", "count"), Carga_total_min=("Carga total (min)", "sum"))
        .rename(columns={"Recurso permitido": "Recurso"})
    )


# ═══════════════════════════════════════════════
# 5. CAPACIDADE LÍQUIDA POR RECURSO
# ═══════════════════════════════════════════════
def calcular_capacidade_por_recurso(calendario: pd.DataFrame) -> pd.DataFrame:
    """Soma a capacidade líquida do Calendário por recurso."""
    return (
        calendario
        .groupby("Recurso", as_index=False)
        .agg(Capacidade_liquida_min=("Capacidade líquida", "sum"))
    )


# ═══════════════════════════════════════════════
# 6. RESUMO POR RECURSO
# ═══════════════════════════════════════════════
def montar_resumo(carga: pd.DataFrame, capacidade: pd.DataFrame) -> pd.DataFrame:
    """
    Une carga e capacidade, calcula utilização, folga/sobrecarga e status.
    """
    resumo = carga.merge(capacidade, on="Recurso", how="outer").fillna(0)

    # Utilização
    resumo["Utilização (%)"] = (
        resumo["Carga_total_min"] / resumo["Capacidade_liquida_min"].replace(0, float("nan")) * 100
    ).round(1)

    # Folga ou excesso
    resumo["Folga (min)"] = (resumo["Capacidade_liquida_min"] - resumo["Carga_total_min"]).round(1)

    # Status
    resumo["Status"] = resumo.apply(
        lambda r: "Sobrecarga" if r["Carga_total_min"] > r["Capacidade_liquida_min"] else "OK",
        axis=1,
    )

    return resumo.sort_values("Utilização (%)", ascending=False).reset_index(drop=True)


# ═══════════════════════════════════════════════
# 7. GARGALO PRINCIPAL
# ═══════════════════════════════════════════════
def identificar_gargalo(resumo: pd.DataFrame) -> str | None:
    """Retorna o nome do recurso com maior utilização."""
    if resumo.empty:
        return None
    idx = resumo["Utilização (%)"].idxmax()
    return resumo.loc[idx, "Recurso"]


# ═══════════════════════════════════════════════
# 8. GRÁFICOS
# ═══════════════════════════════════════════════
def grafico_carga_capacidade(resumo: pd.DataFrame) -> go.Figure:
    """Gráfico de barras duplas: carga vs capacidade por recurso."""
    recursos = resumo["Recurso"]

    fig = go.Figure()
    fig.add_bar(
        name="Capacidade líquida",
        x=recursos,
        y=resumo["Capacidade_liquida_min"],
        marker_color=COR_CAPACIDADE,
        opacity=0.85,
    )
    fig.add_bar(
        name="Carga total",
        x=recursos,
        y=resumo["Carga_total_min"],
        marker_color=[COR_SOBRECARGA if s == "Sobrecarga" else COR_CARGA for s in resumo["Status"]],
        opacity=0.9,
    )

    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        barmode="group",
        title="Carga × Capacidade por Recurso (min)",
        xaxis_title="Recurso",
        yaxis_title="Minutos",
        legend=dict(orientation="h", y=1.08, x=0),
        paper_bgcolor="#0d1117",
        plot_bgcolor="#0d1117",
        font=dict(color="#e6edf3", family="Inter"),
        height=400,
    )
    return fig


def grafico_utilizacao(resumo: pd.DataFrame) -> go.Figure:
    """Gráfico de barras horizontais de utilização (%) com linha em 100 %."""
    cores = [COR_SOBRECARGA if s == "Sobrecarga" else COR_OK for s in resumo["Status"]]

    fig = go.Figure()
    fig.add_bar(
        x=resumo["Utilização (%)"],
        y=resumo["Recurso"],
        orientation="h",
        marker_color=cores,
        text=[f"{v:.1f}%" for v in resumo["Utilização (%)"]],
        textposition="outside",
    )
    # Linha de 100 %
    fig.add_vline(x=100, line_dash="dash", line_color="#8b949e", annotation_text="100 %")

    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        title="Utilização por Recurso (%)",
        xaxis_title="Utilização (%)",
        yaxis=dict(autorange="reversed"),
        paper_bgcolor="#0d1117",
        plot_bgcolor="#0d1117",
        font=dict(color="#e6edf3", family="Inter"),
        height=420,
        margin=dict(r=80),
    )
    return fig


def grafico_pizza_status(resumo: pd.DataFrame) -> go.Figure:
    """Distribuição OK vs Sobrecarga."""
    contagem = resumo["Status"].value_counts().reset_index()
    contagem.columns = ["Status", "Qtd"]

    cores_map = {"OK": COR_OK, "Sobrecarga": COR_SOBRECARGA}
    cores = [cores_map.get(s, "#8b949e") for s in contagem["Status"]]

    fig = go.Figure(go.Pie(
        labels=contagem["Status"],
        values=contagem["Qtd"],
        hole=0.5,
        marker_colors=cores,
        textinfo="label+value",
    ))
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        title="Distribuição de Status",
        paper_bgcolor="#0d1117",
        font=dict(color="#e6edf3", family="Inter"),
        height=340,
        showlegend=False,
    )
    return fig


# ═══════════════════════════════════════════════
# 9. EXPORTAÇÃO DO EXCEL PROCESSADO
# ═══════════════════════════════════════════════
def gerar_excel_saida(enriquecido: pd.DataFrame, resumo: pd.DataFrame, sem_rota: pd.DataFrame) -> bytes:
    """Gera Excel com três abas e retorna bytes para download."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        enriquecido.to_excel(writer, sheet_name="Base Enriquecida", index=False)
        resumo.to_excel(writer, sheet_name="Resumo Recursos", index=False)
        sem_rota.to_excel(writer, sheet_name="Ordens sem Rota", index=False)

        # Formatação básica
        workbook = writer.book
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#1f6feb", "font_color": "#ffffff",
            "border": 1, "align": "center",
        })
        for sheet_name, df in [
            ("Base Enriquecida", enriquecido),
            ("Resumo Recursos", resumo),
            ("Ordens sem Rota", sem_rota),
        ]:
            ws = writer.sheets[sheet_name]
            for col_num, col_name in enumerate(df.columns):
                ws.write(0, col_num, col_name, header_fmt)
                ws.set_column(col_num, col_num, max(len(col_name) + 4, 14))
    return buf.getvalue()


# ═══════════════════════════════════════════════
# 10. HELPERS DE DISPLAY
# ═══════════════════════════════════════════════
def chip_status(status: str) -> str:
    cls = "chip-ok" if status == "OK" else "chip-sobrecarga"
    return f'<span class="{cls}">{status}</span>'


def formatar_resumo_display(resumo: pd.DataFrame) -> pd.DataFrame:
    """Retorna cópia do resumo com colunas renomeadas para exibição."""
    display = resumo.copy()
    display = display.rename(columns={
        "Carga_total_min": "Carga total (min)",
        "Capacidade_liquida_min": "Capacidade líquida (min)",
    })
    return display


# ═══════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🏭 Simulador de Produção")
    st.markdown("**v1.0 MVP** — Laser & Hotstamp")
    st.divider()

    uploaded = st.file_uploader(
        "📂 Carregar planilha Excel",
        type=["xlsx", "xls"],
        help="Planilha com as abas: Recursos, Produtos, Roteiros, Demandas, Calendario, Programacao, Dashboard",
    )
    st.divider()
    st.markdown(
        "<small style='color:#8b949e'>Carga (min) = Qtd × Ciclo(s) ÷ 60 ÷ Rendimento + Setup<br>"
        "Utilização = Carga ÷ Capacidade líquida × 100</small>",
        unsafe_allow_html=True,
    )


# ═══════════════════════════════════════════════
# TELA INICIAL (sem upload)
# ═══════════════════════════════════════════════
if uploaded is None:
    st.markdown("# 🏭 Simulador de Capacidade e Programação")
    st.markdown("### Fábrica — Corte a Laser & Prensa Hotstamp")
    st.divider()

    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("**1. Upload**\nCarregue a planilha Excel com as 7 abas obrigatórias.", icon="📂")
    with col2:
        st.info("**2. Validação**\nO sistema verifica estrutura, abas e colunas automaticamente.", icon="✅")
    with col3:
        st.info("**3. Dashboard**\nVisualize carga, gargalo, utilização e exporte os resultados.", icon="📊")

    st.divider()
    st.markdown("#### Abas obrigatórias na planilha:")
    cols = st.columns(4)
    abas = list(ABAS_OBRIGATORIAS.keys())
    for i, aba in enumerate(abas):
        with cols[i % 4]:
            with st.expander(f"📋 {aba}"):
                for col in ABAS_OBRIGATORIAS[aba]:
                    st.markdown(f"- `{col}`")
    st.stop()


# ═══════════════════════════════════════════════
# PROCESSAMENTO PRINCIPAL
# ═══════════════════════════════════════════════
try:
    xls = pd.ExcelFile(uploaded)
except Exception as e:
    st.error(f"Erro ao abrir o arquivo: {e}")
    st.stop()

# --- Validação ---
erros = validar_excel(xls)
if erros:
    st.error("### ⚠️ Arquivo inválido — corrija os problemas abaixo:")
    for e in erros:
        st.markdown(e)
    st.stop()

st.success("✅ Arquivo validado com sucesso!", icon="✅")

# --- Carregamento ---
dados = carregar_abas(xls)
demandas   = dados["Demandas"]
roteiros   = dados["Roteiros"]
calendario = dados["Calendario"]
recursos   = dados["Recursos"]
programacao= dados["Programacao"]

# --- Processamento ---
enriquecido, sem_rota  = enriquecer_demandas(demandas, roteiros)
carga_rec              = calcular_carga_por_recurso(enriquecido)
capacidade_rec         = calcular_capacidade_por_recurso(calendario)
resumo                 = montar_resumo(carga_rec, capacidade_rec)
gargalo                = identificar_gargalo(resumo)


# ═══════════════════════════════════════════════
# LAYOUT DO DASHBOARD
# ═══════════════════════════════════════════════
st.markdown("# 🏭 Dashboard Executivo de Produção")
st.divider()

# ── KPIs principais ───────────────────────────
carga_total_geral  = resumo["Carga_total_min"].sum()
cap_total_geral    = resumo["Capacidade_liquida_min"].sum()
util_geral         = (carga_total_geral / cap_total_geral * 100) if cap_total_geral else 0
n_sobrecarga       = (resumo["Status"] == "Sobrecarga").sum()
n_ordens           = len(demandas)
n_sem_rota         = len(sem_rota)

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("📦 Total de Ordens",    f"{n_ordens}")
k2.metric("✅ Com Rota",           f"{len(enriquecido)}")
k3.metric("❌ Sem Rota",           f"{n_sem_rota}")
k4.metric("⚡ Carga Total (h)",    f"{carga_total_geral/60:.1f}")
k5.metric("🏋️ Capacidade Total (h)", f"{cap_total_geral/60:.1f}")
k6.metric("📊 Utilização Geral",   f"{util_geral:.1f}%",
          delta="Atenção!" if util_geral > 85 else "Normal",
          delta_color="inverse" if util_geral > 100 else "normal")

st.divider()

# ── Gargalo ───────────────────────────────────
if gargalo:
    gargalo_row = resumo[resumo["Recurso"] == gargalo].iloc[0]
    g_util = gargalo_row["Utilização (%)"]
    cor = COR_SOBRECARGA if g_util > 100 else "#d29922"
    st.markdown(
        f"""
        <div style='background:#161b22;border:1px solid {cor};border-radius:8px;
                    padding:12px 20px;margin-bottom:16px;'>
          <span style='color:{cor};font-size:1rem;font-weight:700;'>🔴 Gargalo Principal:</span>
          <span style='color:#f0f6fc;font-size:1rem;margin-left:8px;'><b>{gargalo}</b></span>
          <span style='color:#8b949e;margin-left:16px;'>Utilização:</span>
          <span style='color:{cor};font-weight:700;margin-left:4px;'>{g_util:.1f}%</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ── Gráficos ──────────────────────────────────
col_g1, col_g2 = st.columns([3, 1])
with col_g1:
    st.plotly_chart(grafico_carga_capacidade(resumo), use_container_width=True)
with col_g2:
    st.plotly_chart(grafico_pizza_status(resumo), use_container_width=True)

st.plotly_chart(grafico_utilizacao(resumo), use_container_width=True)

st.divider()

# ── Tabela: Resumo por Recurso ─────────────────
st.markdown("## 📋 Resumo por Recurso")
display_resumo = formatar_resumo_display(resumo)

# Colorir Status com HTML
html_rows = []
for _, row in display_resumo.iterrows():
    status_html = chip_status(row["Status"])
    html_rows.append({
        "Recurso": row["Recurso"],
        "Ordens": int(row["Ordens"]) if "Ordens" in row else "—",
        "Carga total (min)": f"{row['Carga total (min)']:.1f}",
        "Capacidade líquida (min)": f"{row['Capacidade líquida (min)']:.1f}",
        "Folga (min)": f"{row['Folga (min)']:.1f}",
        "Utilização (%)": f"{row['Utilização (%)']:.1f}%",
        "Status": row["Status"],
    })

df_display = pd.DataFrame(html_rows)
st.dataframe(
    df_display,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Status": st.column_config.TextColumn("Status"),
        "Utilização (%)": st.column_config.TextColumn("Utilização (%)"),
    },
)

st.divider()

# ── Tabela: Ordens Enriquecidas ───────────────
st.markdown("## 🗂️ Ordens com Rota Válida")
cols_show = [
    "Ordem", "Produto", "Cliente", "Quantidade", "Data de necessidade",
    "Prioridade", "Destino", "Status",
    "Recurso permitido", "Tempo de ciclo (s)", "Setup (min)", "Rendimento",
    "Carga produção (min)", "Carga total (min)",
]
cols_show = [c for c in cols_show if c in enriquecido.columns]

st.dataframe(
    enriquecido[cols_show].reset_index(drop=True),
    use_container_width=True,
    hide_index=True,
)

st.divider()

# ── Tabela: Ordens sem Rota ───────────────────
st.markdown("## ⚠️ Ordens sem Rota Válida")
if sem_rota.empty:
    st.success("Nenhuma ordem sem rota. Todos os produtos têm roteiro definido.", icon="✅")
else:
    st.warning(f"{len(sem_rota)} ordem(ns) sem rota definida no roteiro.", icon="⚠️")
    cols_sem_rota = [c for c in ["Ordem", "Produto", "Cliente", "Quantidade", "Data de necessidade", "Prioridade"] if c in sem_rota.columns]
    st.dataframe(sem_rota[cols_sem_rota].reset_index(drop=True), use_container_width=True, hide_index=True)

st.divider()

# ── Expansores de dados brutos ─────────────────
with st.expander("🔍 Ver dados brutos — Calendário"):
    st.dataframe(calendario, use_container_width=True, hide_index=True)

with st.expander("🔍 Ver dados brutos — Recursos"):
    st.dataframe(recursos, use_container_width=True, hide_index=True)

with st.expander("🔍 Ver dados brutos — Programação"):
    st.dataframe(programacao, use_container_width=True, hide_index=True)

st.divider()

# ── Exportação ────────────────────────────────
st.markdown("## 💾 Exportar Resultado")
excel_bytes = gerar_excel_saida(enriquecido, display_resumo, sem_rota)
st.download_button(
    label="⬇️  Baixar Excel Processado",
    data=excel_bytes,
    file_name="simulador_producao_resultado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.markdown(
    "<br><small style='color:#8b949e'>Simulador de Produção · MVP v1.0 · "
    "Corte a Laser & Hotstamp · Desenvolvido com Streamlit + Plotly</small>",
    unsafe_allow_html=True,
)
