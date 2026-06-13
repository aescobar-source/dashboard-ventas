"""
Dashboard Ejecutivo de Ventas 2026
Consultor Senior BI — Paleta Gris Elegante
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
#  CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Ejecutivo de Ventas 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  PALETA GRIS ELEGANTE
# ─────────────────────────────────────────────
COLORS = {
    "bg":           "#F8F8F8",
    "card":         "#FFFFFF",
    "border":       "#E0E0E0",
    "text_dark":    "#1A1A1A",
    "text_mid":     "#555555",
    "text_light":   "#888888",
    "accent":       "#2C2C2C",
    "gray1":        "#E8E8E8",
    "gray2":        "#C0C0C0",
    "gray3":        "#8A8A8A",
    "gray4":        "#4A4A4A",
    "gray5":        "#1A1A1A",
    "pos":          "#3A3A3A",
    "neg":          "#A0A0A0",
    "highlight":    "#222222",
}

GRAY_SCALE = ["#E8E8E8", "#C0C0C0", "#9A9A9A", "#6A6A6A", "#3A3A3A",
              "#2A2A2A", "#1A1A1A"]

# ─────────────────────────────────────────────
#  HELPER: convierte hex a rgba (Plotly no acepta hex+alpha)
# ─────────────────────────────────────────────
def hex_to_rgba(hex_color, alpha=0.15):
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"



# ─────────────────────────────────────────────
#  CSS GLOBAL
# ─────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

  html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    background-color: {COLORS['bg']};
    color: {COLORS['text_dark']};
  }}
  .stApp {{ background-color: {COLORS['bg']}; }}

  /* Sidebar */
  section[data-testid="stSidebar"] {{
    background-color: {COLORS['accent']};
  }}
  section[data-testid="stSidebar"] * {{
    color: #EEEEEE !important;
  }}
  section[data-testid="stSidebar"] .stSelectbox label,
  section[data-testid="stSidebar"] .stMultiSelect label {{
    color: #CCCCCC !important;
    font-size: 0.78rem;
    font-weight: 500;
    letter-spacing: 0.06em;
    text-transform: uppercase;
  }}

  /* KPI Cards */
  .kpi-card {{
    background: {COLORS['card']};
    border: 1px solid {COLORS['border']};
    border-radius: 8px;
    padding: 20px 24px;
    margin-bottom: 12px;
  }}
  .kpi-label {{
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.10em;
    text-transform: uppercase;
    color: {COLORS['text_light']};
    margin-bottom: 6px;
  }}
  .kpi-value {{
    font-size: 1.85rem;
    font-weight: 700;
    color: {COLORS['text_dark']};
    line-height: 1.1;
  }}
  .kpi-delta-pos {{
    font-size: 0.78rem;
    color: #333333;
    font-weight: 500;
    margin-top: 4px;
  }}
  .kpi-delta-neg {{
    font-size: 0.78rem;
    color: #888888;
    font-weight: 500;
    margin-top: 4px;
  }}
  .kpi-sub {{
    font-size: 0.72rem;
    color: {COLORS['text_light']};
    margin-top: 2px;
  }}

  /* Section Headers */
  .section-header {{
    font-size: 0.70rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: {COLORS['text_light']};
    border-bottom: 1px solid {COLORS['border']};
    padding-bottom: 8px;
    margin-bottom: 16px;
    margin-top: 32px;
  }}
  .page-title {{
    font-size: 1.6rem;
    font-weight: 700;
    color: {COLORS['text_dark']};
    letter-spacing: -0.02em;
  }}
  .page-subtitle {{
    font-size: 0.85rem;
    color: {COLORS['text_light']};
    margin-top: 2px;
  }}

  /* Tab strip */
  .stTabs [data-baseweb="tab-list"] {{
    gap: 4px;
    background: transparent;
    border-bottom: 1px solid {COLORS['border']};
  }}
  .stTabs [data-baseweb="tab"] {{
    font-size: 0.78rem;
    font-weight: 500;
    letter-spacing: 0.04em;
    color: {COLORS['text_light']};
    background: transparent;
    border: none;
    padding: 8px 16px;
  }}
  .stTabs [aria-selected="true"] {{
    color: {COLORS['text_dark']} !important;
    border-bottom: 2px solid {COLORS['text_dark']} !important;
    background: transparent !important;
  }}

  /* Charts background */
  .js-plotly-plot .plotly {{ background: {COLORS['card']} !important; }}

  /* Hide default streamlit elements */
  #MainMenu, footer, header {{ visibility: hidden; }}
  .block-container {{ padding-top: 1.5rem; padding-bottom: 2rem; }}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  CARGA Y LIMPIEZA DE DATOS
# ─────────────────────────────────────────────
@st.cache_data
def load_data():
    # ESTRATEGIA DE ESCALABILIDAD:
    # El archivo es la única fuente de verdad. Al añadir filas con
    # los meses junio-diciembre 2026, el dashboard se actualiza solo.
    # Recomendación: mantener el mismo esquema de columnas.
    df = pd.read_excel("Ventas_enero-mayo_2026.xlsx")
    df.columns = df.columns.str.strip()

    # Limpieza de strings (astype(object) para manejar NaN en ArrowStringArray)
    for col in ["TIPO DE VENTA", "ZONA", "CIUDAD", "SEGMENTO",
                "MARCA", "TIPO DE PRODUCTO", "EJECUTVO", "MES"]:
        if col in df.columns:
            df[col] = (df[col].astype(object)
                              .fillna("SIN DATO")
                              .astype(str)
                              .str.strip()
                              .str.upper())

    # Unificar variantes de la misma categoría
    df["TIPO DE PRODUCTO"] = df["TIPO DE PRODUCTO"].str.replace(r"\s+", " ", regex=True).str.strip()
    df["CIUDAD"] = df["CIUDAD"].str.replace(r"\s+", " ", regex=True).str.strip()

    # ADIDAS / adidas → ADIDAS
    df["MARCA"] = df["MARCA"].str.upper()

    # Orden cronológico de meses
    MES_ORDER = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
                 "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
                 "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12}
    df["MES_NUM"] = df["MES"].map(MES_ORDER).fillna(99).astype(int)
    df = df.sort_values("MES_NUM")

    # Márgenes derivados
    df["MARGEN_%"] = np.where(
        df["TOTAL VENTA"] > 0,
        df["UTILIDAD TOTAL $"] / df["TOTAL VENTA"],
        0
    )
    df["COSTO_TOTAL"] = df["COSTO BODEGA TOTAL"].fillna(0)

    # Efectividad de descuento: % de filas con descuento > 0
    df["TIENE_DESCUENTO"] = df["TOTAL DESCUENTO"] > 0

    # Ticket promedio por factura
    df["FACTURA_KEY"] = df["FACTURA"].fillna(pd.Series(df.index.astype(str), index=df.index))

    return df

df = load_data()

# ─────────────────────────────────────────────
#  COORDENADAS DE CIUDADES (Nicaragua)
# ─────────────────────────────────────────────
CITY_COORDS = {
    "MANAGUA":               (12.1364, -86.2914),
    "LEON":                  (12.4372, -86.8796),
    "MATAGALPA":             (12.9269, -85.9183),
    "ESTELI":                (13.0844, -86.3609),
    "CHINANDEGA":            (12.6253, -87.1253),
    "MASAYA":                (11.9745, -86.0955),
    "JINOTEGA":              (13.0958, -85.9989),
    "JUIGALPA":              (12.1024, -85.3681),
    "OCOTAL":                (13.6348, -86.4748),
    "SOMOTO":                (13.4748, -86.5897),
    "BLUEFIELDS":            (11.9944, -83.7733),
    "NUEVA GUINEA":          (11.6951, -84.4566),
    "RIVAS":                 (11.4368, -85.8369),
    "DIRIAMBA":              (11.8547, -86.2385),
    "BOACO":                 (12.4739, -85.6595),
    "JALAPA":                (13.9253, -86.1217),
    "SIUNA":                 (13.7317, -84.7747),
    "WASLALA":               (13.3217, -85.3603),
    "MATIGUAS":              (12.8342, -85.4585),
    "RIO BLANCO":            (12.9592, -85.2226),
    "CAMOAPA":               (12.3885, -85.5164),
    "LA TRINIDAD":           (12.9718, -86.2384),
    "SEBACO":                (12.8519, -86.1001),
    "CIUDAD DARIO":          (12.7296, -86.1163),
    "CONDEGA":               (13.3617, -86.3983),
    "EL JICARO":             (13.7085, -86.1270),
    "QUILALI":               (13.6524, -86.0224),
    "SAN JUAN DEL RIO COCO": (13.5469, -86.1617),
    "WIWILI":                (13.6194, -85.8181),
    "PANTASMA":              (13.3378, -85.9394),
    "LA DALIA":              (13.1575, -85.7339),
    "EL CUA":                (13.3667, -85.6833),
    "MULUKUKU":              (13.1500, -84.9667),
    "YALI":                  (13.3000, -86.0000),
    "LA CONCORDIA":          (13.1897, -86.1737),
    "SAN RAFAEL DEL NORTE":  (13.2099, -86.1092),
    "MALPASILLO":            (12.6167, -86.6000),
    "EL SAUCE":              (12.8931, -86.5364),
    "CHICHIGALPA":           (12.5642, -87.0280),
    "DIRIA":                 (11.8736, -86.0681),
    "NUEVA SEGOVIA":         (13.6348, -86.4748),
    "EL RAMA":               (12.1569, -84.2186),
    "SANTO TOMAS":           (12.0683, -84.9383),
    "SUSUCAYAN":             (14.0128, -86.1981),
    "PBMO":                  (12.1364, -86.2914),
}

# ─────────────────────────────────────────────
#  SIDEBAR — FILTROS
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding: 16px 0 20px; border-bottom: 1px solid #444;'>
      <div style='font-size:1.1rem; font-weight:700; color:#FFF; letter-spacing:-0.01em;'>
        📊 Ventas 2026
      </div>
      <div style='font-size:0.72rem; color:#AAA; margin-top:3px;'>Dashboard Ejecutivo</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    meses_disp = sorted(df["MES"].unique(),
                        key=lambda x: {"ENERO":1,"FEBRERO":2,"MARZO":3,
                                       "ABRIL":4,"MAYO":5,"JUNIO":6,
                                       "JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,
                                       "OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12}.get(x, 99))

    sel_mes   = st.multiselect("Mes",          meses_disp,          default=meses_disp)
    sel_zona  = st.multiselect("Zona",         sorted(df["ZONA"].unique()),   default=sorted(df["ZONA"].unique()))
    sel_seg   = st.multiselect("Segmento",     sorted(df["SEGMENTO"].unique()), default=sorted(df["SEGMENTO"].unique()))
    sel_marca = st.multiselect("Marca",        sorted(df["MARCA"].unique()),   default=sorted(df["MARCA"].unique()))
    sel_prod  = st.multiselect("Tipo Producto",sorted(df["TIPO DE PRODUCTO"].unique()),
                               default=sorted(df["TIPO DE PRODUCTO"].unique()))
    sel_ejec  = st.multiselect("Ejecutivo",    sorted(df["EJECUTVO"].unique()), default=sorted(df["EJECUTVO"].unique()))

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:0.68rem;color:#777;'>Actualiza los filtros para análisis dinámico.</div>",
                unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  FILTRAR DATOS
# ─────────────────────────────────────────────
mask = (
    df["MES"].isin(sel_mes) &
    df["ZONA"].isin(sel_zona) &
    df["SEGMENTO"].isin(sel_seg) &
    df["MARCA"].isin(sel_marca) &
    df["TIPO DE PRODUCTO"].isin(sel_prod) &
    df["EJECUTVO"].isin(sel_ejec)
)
dff = df[mask].copy()

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def fmt_usd(v):
    if abs(v) >= 1_000_000:
        return f"${v/1_000_000:.2f}M"
    elif abs(v) >= 1_000:
        return f"${v/1_000:.1f}K"
    return f"${v:,.0f}"

def fmt_pct(v):
    return f"{v*100:.1f}%"

def kpi_card(label, value, delta_text="", delta_pos=True, sub=""):
    delta_cls = "kpi-delta-pos" if delta_pos else "kpi-delta-neg"
    delta_icon = "▲" if delta_pos else "▼"
    delta_html = f'<div class="{delta_cls}">{delta_icon} {delta_text}</div>' if delta_text else ""
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return f"""
    <div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>
      {delta_html}
      {sub_html}
    </div>
    """

def chart_layout(fig, title="", height=380):
    fig.update_layout(
        title=dict(text=title, font=dict(size=13, color=COLORS["text_dark"],
                   family="Inter"), x=0, xanchor="left"),
        plot_bgcolor=COLORS["card"],
        paper_bgcolor=COLORS["card"],
        font=dict(family="Inter", color=COLORS["text_mid"], size=11),
        height=height,
        margin=dict(l=10, r=10, t=40, b=10),
        legend=dict(bgcolor="rgba(0,0,0,0)", borderwidth=0,
                    font=dict(size=10)),
        xaxis=dict(gridcolor=COLORS["gray1"], linecolor=COLORS["border"],
                   tickfont=dict(size=10)),
        yaxis=dict(gridcolor=COLORS["gray1"], linecolor=COLORS["border"],
                   tickfont=dict(size=10)),
    )
    return fig

# ─────────────────────────────────────────────
#  CÁLCULO DE KPIs PRINCIPALES
# ─────────────────────────────────────────────
total_venta       = dff["TOTAL VENTA"].sum()
total_utilidad    = dff["UTILIDAD TOTAL $"].sum()
total_descuento   = dff["TOTAL DESCUENTO"].sum()
total_costo       = dff["COSTO_TOTAL"].sum()
total_pares       = dff["PARES"].sum()
total_transacc    = dff["FACTURA_KEY"].nunique()
margen_global     = total_utilidad / total_venta if total_venta > 0 else 0
ticket_prom       = total_venta / total_transacc if total_transacc > 0 else 0
descuento_rate    = total_descuento / (dff["SUBTOTAL"].sum()) if dff["SUBTOTAL"].sum() > 0 else 0
efect_desc        = dff["TIENE_DESCUENTO"].mean()

# Mes con mejor utilidad
mes_util = dff.groupby("MES")["UTILIDAD TOTAL $"].sum()
best_mes  = mes_util.idxmax() if not mes_util.empty else "—"

# ─────────────────────────────────────────────
#  HEADER
# ─────────────────────────────────────────────
col_title, col_info = st.columns([3, 1])
with col_title:
    st.markdown("""
    <div class="page-title">Dashboard Ejecutivo de Ventas</div>
    <div class="page-subtitle">Enero – Mayo 2026 &nbsp;·&nbsp; Actualización mensual automática</div>
    """, unsafe_allow_html=True)
with col_info:
    st.markdown(f"""
    <div style='text-align:right; padding-top:8px'>
      <div style='font-size:0.72rem;color:{COLORS["text_light"]};letter-spacing:0.06em;'>REGISTROS ACTIVOS</div>
      <div style='font-size:1.4rem;font-weight:700;color:{COLORS["text_dark"]};'>{len(dff):,}</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  KPIs — FILA 1
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">Indicadores Clave de Rendimiento</div>', unsafe_allow_html=True)

c1, c2, c3, c4, c5, c6 = st.columns(6)
with c1:
    st.markdown(kpi_card("Venta Total", fmt_usd(total_venta),
                sub=f"{total_transacc:,} facturas"), unsafe_allow_html=True)
with c2:
    st.markdown(kpi_card("Utilidad Bruta", fmt_usd(total_utilidad),
                sub=f"Margen: {fmt_pct(margen_global)}"), unsafe_allow_html=True)
with c3:
    st.markdown(kpi_card("Margen %", fmt_pct(margen_global),
                delta_text=f"Mejor mes: {best_mes.title()}", delta_pos=(margen_global > 0.25),
                sub="Meta sugerida > 25%"), unsafe_allow_html=True)
with c4:
    st.markdown(kpi_card("Ticket Promedio", fmt_usd(ticket_prom),
                sub=f"{total_pares:,} unidades vendidas"), unsafe_allow_html=True)
with c5:
    st.markdown(kpi_card("Total Descuentos", fmt_usd(total_descuento),
                delta_text=fmt_pct(descuento_rate) + " sobre subtotal", delta_pos=False,
                sub=f"{fmt_pct(efect_desc)} transacc. con dcto."), unsafe_allow_html=True)
with c6:
    st.markdown(kpi_card("Costo Total", fmt_usd(total_costo),
                sub=f"Costo/Venta: {fmt_pct(total_costo/total_venta) if total_venta > 0 else '0%'}"),
                unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  TABS PRINCIPALES
# ─────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "  Tendencias  ", "  Rentabilidad  ", "  Mapa Geoespacial  ",
    "  Ejecutivos  ", "  Marcas & Productos  ", "  Análisis de Variaciones  "
])

# ══════════════════════════════════════════════
#  TAB 1 — TENDENCIAS MENSUALES
# ══════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-header">Evolución Mensual</div>', unsafe_allow_html=True)

    mes_gb = dff.groupby("MES").agg(
        Venta=("TOTAL VENTA", "sum"),
        Utilidad=("UTILIDAD TOTAL $", "sum"),
        Descuento=("TOTAL DESCUENTO", "sum"),
        Costo=("COSTO_TOTAL", "sum"),
        Pares=("PARES", "sum"),
        Facturas=("FACTURA_KEY", "nunique"),
    ).reset_index()
    mes_gb["Margen_%"] = mes_gb["Utilidad"] / mes_gb["Venta"].replace(0, np.nan)
    mes_gb["MES_NUM"] = mes_gb["MES"].map(
        {"ENERO":1,"FEBRERO":2,"MARZO":3,"ABRIL":4,"MAYO":5,"JUNIO":6,
         "JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,"OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12})
    mes_gb = mes_gb.sort_values("MES_NUM")
    mes_labels = mes_gb["MES"].str.capitalize().tolist()

    col_left, col_right = st.columns([2, 1])

    with col_left:
        # Venta + Utilidad + Descuento
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(
            x=mes_labels, y=mes_gb["Venta"], name="Venta Total",
            marker_color=COLORS["gray4"], opacity=0.85,
        ), secondary_y=False)
        fig.add_trace(go.Bar(
            x=mes_labels, y=mes_gb["Utilidad"], name="Utilidad Bruta",
            marker_color=COLORS["gray3"], opacity=0.85,
        ), secondary_y=False)
        fig.add_trace(go.Scatter(
            x=mes_labels, y=mes_gb["Margen_%"] * 100, name="Margen %",
            mode="lines+markers+text",
            line=dict(color=COLORS["gray5"], width=2.5, dash="solid"),
            marker=dict(size=7, color=COLORS["gray5"]),
            text=[f"{v:.1f}%" for v in mes_gb["Margen_%"].fillna(0) * 100],
            textposition="top center", textfont=dict(size=10),
        ), secondary_y=True)
        fig.update_yaxes(title_text="USD $", secondary_y=False,
                         tickprefix="$", gridcolor=COLORS["gray1"])
        fig.update_yaxes(title_text="Margen %", secondary_y=True,
                         ticksuffix="%", showgrid=False)
        chart_layout(fig, "Venta · Utilidad · Margen %", height=350)
        fig.update_layout(barmode="group", legend=dict(orientation="h", y=-0.15))
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        # Descuentos mensuales
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=mes_labels, y=mes_gb["Descuento"],
            marker_color=[COLORS["gray2"] if v < mes_gb["Descuento"].mean()
                          else COLORS["gray4"] for v in mes_gb["Descuento"]],
            text=[fmt_usd(v) for v in mes_gb["Descuento"]],
            textposition="outside", textfont=dict(size=10),
        ))
        chart_layout(fig2, "Descuentos por Mes", height=350)
        fig2.update_layout(showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    # Línea de tiempo: Volumen de unidades
    col_a, col_b = st.columns(2)
    with col_a:
        fig3 = go.Figure()
        fig3.add_trace(go.Scatter(
            x=mes_labels, y=mes_gb["Pares"],
            fill="tozeroy", mode="lines+markers",
            line=dict(color=COLORS["gray4"], width=2.5),
            fillcolor="rgba(74,74,74,0.12)",
            marker=dict(size=7),
        ))
        chart_layout(fig3, "Volumen de Unidades Vendidas (Pares/Items)", height=300)
        st.plotly_chart(fig3, use_container_width=True)

    with col_b:
        fig4 = go.Figure()
        fig4.add_trace(go.Scatter(
            x=mes_labels, y=mes_gb["Facturas"],
            mode="lines+markers",
            line=dict(color=COLORS["gray3"], width=2.5, dash="dot"),
            marker=dict(size=7, symbol="diamond"),
        ))
        chart_layout(fig4, "Número de Facturas Emitidas por Mes", height=300)
        st.plotly_chart(fig4, use_container_width=True)

    # Segmento por mes
    st.markdown('<div class="section-header">Venta por Segmento × Mes</div>', unsafe_allow_html=True)
    seg_mes = dff.groupby(["MES", "SEGMENTO"])["TOTAL VENTA"].sum().reset_index()
    seg_mes["MES_NUM"] = seg_mes["MES"].map(
        {"ENERO":1,"FEBRERO":2,"MARZO":3,"ABRIL":4,"MAYO":5,"JUNIO":6,
         "JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,"OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12})
    seg_mes = seg_mes.sort_values("MES_NUM")
    seg_mes["MES_LABEL"] = seg_mes["MES"].str.capitalize()

    segs = seg_mes["SEGMENTO"].unique()
    fig5 = go.Figure()
    for i, seg in enumerate(segs):
        s = seg_mes[seg_mes["SEGMENTO"] == seg]
        fig5.add_trace(go.Bar(
            x=s["MES_LABEL"], y=s["TOTAL VENTA"],
            name=seg.title(),
            marker_color=GRAY_SCALE[i % len(GRAY_SCALE)],
        ))
    chart_layout(fig5, "Composición de Venta por Segmento", height=320)
    fig5.update_layout(barmode="stack", legend=dict(orientation="h", y=-0.2))
    st.plotly_chart(fig5, use_container_width=True)


# ══════════════════════════════════════════════
#  TAB 2 — RENTABILIDAD
# ══════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-header">Análisis de Rentabilidad</div>', unsafe_allow_html=True)

    # Margen por segmento
    seg_rent = dff.groupby("SEGMENTO").agg(
        Venta=("TOTAL VENTA","sum"),
        Utilidad=("UTILIDAD TOTAL $","sum"),
        Descuento=("TOTAL DESCUENTO","sum"),
        Costo=("COSTO_TOTAL","sum"),
    ).reset_index()
    seg_rent["Margen%"] = seg_rent["Utilidad"] / seg_rent["Venta"].replace(0, np.nan) * 100
    seg_rent["Desc%"] = seg_rent["Descuento"] / (seg_rent["Venta"] + seg_rent["Descuento"]).replace(0,np.nan) * 100

    col1, col2 = st.columns(2)
    with col1:
        fig = go.Figure()
        for i, row in seg_rent.iterrows():
            fig.add_trace(go.Bar(
                x=[row["SEGMENTO"].title()],
                y=[row["Margen%"]],
                name=row["SEGMENTO"].title(),
                marker_color=GRAY_SCALE[i],
                text=[f"{row['Margen%']:.1f}%"],
                textposition="outside",
            ))
        chart_layout(fig, "Margen % por Segmento", height=320)
        fig.update_layout(showlegend=False, yaxis_ticksuffix="%")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Waterfall: de Venta a Utilidad
        fig_wf = go.Figure(go.Waterfall(
            name="Desglose", orientation="v",
            x=["Venta Bruta", "(-) Descuentos", "(-) Costo Mercancía", "= Utilidad Bruta"],
            y=[total_venta, -total_descuento,
               -total_costo, total_utilidad],
            connector={"line": {"color": COLORS["gray2"]}},
            decreasing={"marker": {"color": COLORS["gray2"]}},
            increasing={"marker": {"color": COLORS["gray4"]}},
            totals={"marker": {"color": COLORS["gray5"]}},
            text=[fmt_usd(total_venta), fmt_usd(-total_descuento),
                  fmt_usd(-total_costo), fmt_usd(total_utilidad)],
            textposition="outside",
        ))
        chart_layout(fig_wf, "Waterfall: Venta → Utilidad Bruta", height=320)
        st.plotly_chart(fig_wf, use_container_width=True)

    # Scatter: Venta vs Margen por zona
    zona_scatter = dff.groupby("ZONA").agg(
        Venta=("TOTAL VENTA","sum"),
        Utilidad=("UTILIDAD TOTAL $","sum"),
        Transacc=("FACTURA_KEY","nunique"),
    ).reset_index()
    zona_scatter["Margen%"] = zona_scatter["Utilidad"] / zona_scatter["Venta"].replace(0,np.nan) * 100

    col3, col4 = st.columns(2)
    with col3:
        fig_sc = go.Figure()
        for i, row in zona_scatter.iterrows():
            fig_sc.add_trace(go.Scatter(
                x=[row["Venta"]], y=[row["Margen%"]],
                mode="markers+text",
                marker=dict(size=max(10, row["Transacc"] / 30),
                            color=GRAY_SCALE[i % len(GRAY_SCALE)],
                            line=dict(color="#000", width=1)),
                text=[row["ZONA"].title()],
                textposition="top center",
                name=row["ZONA"].title(),
            ))
        chart_layout(fig_sc, "Zona: Venta Total vs Margen% (tamaño = transacciones)", height=360)
        fig_sc.update_layout(showlegend=False,
                             xaxis_tickprefix="$", yaxis_ticksuffix="%")
        st.plotly_chart(fig_sc, use_container_width=True)

    with col4:
        # Efectividad de descuentos: con vs sin descuento
        disc_g = dff.groupby("TIENE_DESCUENTO").agg(
            Venta=("TOTAL VENTA","sum"),
            Utilidad=("UTILIDAD TOTAL $","sum"),
            Cnt=("FACTURA_KEY","count"),
        ).reset_index()
        disc_g["Label"] = disc_g["TIENE_DESCUENTO"].map({True:"Con Descuento", False:"Sin Descuento"})
        disc_g["Margen%"] = disc_g["Utilidad"] / disc_g["Venta"].replace(0,np.nan) * 100

        fig_d = make_subplots(rows=1, cols=2,
                              subplot_titles=("Venta Total", "Margen %"))
        fig_d.add_trace(go.Bar(
            x=disc_g["Label"], y=disc_g["Venta"],
            marker_color=[COLORS["gray3"], COLORS["gray5"]],
            text=[fmt_usd(v) for v in disc_g["Venta"]],
            textposition="outside",
        ), row=1, col=1)
        fig_d.add_trace(go.Bar(
            x=disc_g["Label"], y=disc_g["Margen%"],
            marker_color=[COLORS["gray3"], COLORS["gray5"]],
            text=[f"{v:.1f}%" for v in disc_g["Margen%"]],
            textposition="outside",
        ), row=1, col=2)
        chart_layout(fig_d, "Efectividad de Descuentos: Con vs Sin", height=360)
        fig_d.update_layout(showlegend=False)
        st.plotly_chart(fig_d, use_container_width=True)

    # Tabla resumen de rentabilidad por zona
    st.markdown('<div class="section-header">Tabla de Rentabilidad por Zona</div>', unsafe_allow_html=True)
    zona_tbl = dff.groupby("ZONA").agg(
        Venta=("TOTAL VENTA","sum"),
        Utilidad=("UTILIDAD TOTAL $","sum"),
        Costo=("COSTO_TOTAL","sum"),
        Descuento=("TOTAL DESCUENTO","sum"),
        Facturas=("FACTURA_KEY","nunique"),
    ).reset_index()
    zona_tbl["Margen%"] = (zona_tbl["Utilidad"] / zona_tbl["Venta"].replace(0,np.nan) * 100).round(1)
    zona_tbl["Venta"]    = zona_tbl["Venta"].apply(fmt_usd)
    zona_tbl["Utilidad"] = zona_tbl["Utilidad"].apply(fmt_usd)
    zona_tbl["Costo"]    = zona_tbl["Costo"].apply(fmt_usd)
    zona_tbl["Descuento"]= zona_tbl["Descuento"].apply(fmt_usd)
    zona_tbl.columns = ["Zona","Venta Total","Utilidad $","Costo Total","Descuento","Facturas","Margen %"]
    zona_tbl["Zona"] = zona_tbl["Zona"].str.title()
    st.dataframe(zona_tbl, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
#  TAB 3 — MAPA GEOESPACIAL
# ══════════════════════════════════════════════
with tab3:
    st.markdown('<div class="section-header">Mapa de Calor Geoespacial — Nicaragua</div>',
                unsafe_allow_html=True)

    metrica_mapa = st.radio(
        "Métrica a visualizar en el mapa:",
        ["Venta Total", "Utilidad Total", "Descuento Total", "Nº de Transacciones"],
        horizontal=True,
    )
    campo_map = {
        "Venta Total": "TOTAL VENTA",
        "Utilidad Total": "UTILIDAD TOTAL $",
        "Descuento Total": "TOTAL DESCUENTO",
        "Nº de Transacciones": "FACTURA_KEY",
    }[metrica_mapa]

    if campo_map == "FACTURA_KEY":
        city_data = dff.groupby("CIUDAD")[campo_map].nunique().reset_index()
        city_data.columns = ["CIUDAD","VALOR"]
    else:
        city_data = dff.groupby("CIUDAD")[campo_map].sum().reset_index()
        city_data.columns = ["CIUDAD","VALOR"]

    city_data["CIUDAD"] = city_data["CIUDAD"].str.strip().str.upper()
    city_data["lat"] = city_data["CIUDAD"].map(lambda c: CITY_COORDS.get(c, (None,None))[0])
    city_data["lon"] = city_data["CIUDAD"].map(lambda c: CITY_COORDS.get(c, (None,None))[1])
    city_data = city_data.dropna(subset=["lat","lon"])
    city_data["Ciudad_Label"] = city_data["CIUDAD"].str.title()

    if campo_map == "FACTURA_KEY":
        city_data["Texto"] = city_data.apply(
            lambda r: f"{r['Ciudad_Label']}<br>Transacciones: {r['VALOR']:,}", axis=1)
    else:
        city_data["Texto"] = city_data.apply(
            lambda r: f"{r['Ciudad_Label']}<br>{metrica_mapa}: {fmt_usd(r['VALOR'])}", axis=1)

    # Mapa de burbujas densidad
    fig_map = go.Figure()

    # Capa de calor (densidad)
    fig_map.add_trace(go.Densitymapbox(
        lat=city_data["lat"],
        lon=city_data["lon"],
        z=city_data["VALOR"],
        radius=45,
        colorscale=[
            [0.0,  "rgba(230,230,230,0.0)"],
            [0.2,  "rgba(180,180,180,0.3)"],
            [0.5,  "rgba(110,110,110,0.6)"],
            [0.8,  "rgba(50,50,50,0.8)"],
            [1.0,  "rgba(10,10,10,1.0)"],
        ],
        opacity=0.7,
        name="Densidad",
        showscale=False,
    ))

    # Burbujas proporcionales
    max_val = city_data["VALOR"].max()
    fig_map.add_trace(go.Scattermapbox(
        lat=city_data["lat"],
        lon=city_data["lon"],
        mode="markers+text",
        marker=dict(
            size=city_data["VALOR"] / max_val * 55 + 8,
            color=city_data["VALOR"],
            colorscale=[
                [0,   "#E8E8E8"],
                [0.3, "#9A9A9A"],
                [0.7, "#4A4A4A"],
                [1,   "#1A1A1A"],
            ],
            opacity=0.85,
            showscale=True,
            colorbar=dict(
                title=dict(text=metrica_mapa, font=dict(size=11, color="#444")),
                tickfont=dict(size=9, color="#444"),
                bgcolor="rgba(255,255,255,0.9)",
                bordercolor="#DDD",
                x=1.02,
            ),
        ),
        text=city_data["Ciudad_Label"],
        textfont=dict(size=9, color="#111"),
        hovertext=city_data["Texto"],
        hoverinfo="text",
        name="Ciudad",
    ))

    fig_map.update_layout(
        mapbox=dict(
            style="carto-positron",
            center=dict(lat=12.8, lon=-85.5),
            zoom=6.0,
        ),
        paper_bgcolor=COLORS["card"],
        height=580,
        margin=dict(l=0, r=0, t=30, b=0),
        title=dict(
            text=f"Nicaragua · {metrica_mapa} por Ciudad",
            font=dict(size=13, color=COLORS["text_dark"]),
            x=0.01,
        ),
        showlegend=False,
    )
    st.plotly_chart(fig_map, use_container_width=True)

    # Top ciudades
    st.markdown('<div class="section-header">Ranking de Ciudades</div>', unsafe_allow_html=True)
    top_city = city_data.sort_values("VALOR", ascending=False).head(15)
    fig_tc = go.Figure(go.Bar(
        x=top_city["VALOR"],
        y=top_city["Ciudad_Label"],
        orientation="h",
        marker_color=[GRAY_SCALE[min(i, len(GRAY_SCALE)-1)] for i in range(len(top_city))],
        text=[fmt_usd(v) if campo_map != "FACTURA_KEY" else f"{v:,}" for v in top_city["VALOR"]],
        textposition="outside",
    ))
    chart_layout(fig_tc, f"Top 15 Ciudades — {metrica_mapa}", height=420)
    fig_tc.update_layout(
        yaxis=dict(autorange="reversed"),
        showlegend=False,
    )
    st.plotly_chart(fig_tc, use_container_width=True)

    # Comparativa por Zona
    zona_map = dff.groupby("ZONA").agg(
        Venta=("TOTAL VENTA","sum"),
        Utilidad=("UTILIDAD TOTAL $","sum"),
    ).reset_index()
    zona_map["Zona_L"] = zona_map["ZONA"].str.title()

    col_z1, col_z2 = st.columns(2)
    with col_z1:
        fig_zp = go.Figure(go.Pie(
            labels=zona_map["Zona_L"], values=zona_map["Venta"],
            hole=0.55,
            marker=dict(colors=GRAY_SCALE[:len(zona_map)],
                        line=dict(color="#FFFFFF", width=2)),
            textfont=dict(size=10),
        ))
        chart_layout(fig_zp, "Participación de Venta por Zona", height=340)
        fig_zp.update_layout(legend=dict(orientation="v", font=dict(size=10)))
        st.plotly_chart(fig_zp, use_container_width=True)

    with col_z2:
        fig_zu = go.Figure(go.Pie(
            labels=zona_map["Zona_L"], values=zona_map["Utilidad"],
            hole=0.55,
            marker=dict(colors=GRAY_SCALE[:len(zona_map)],
                        line=dict(color="#FFFFFF", width=2)),
            textfont=dict(size=10),
        ))
        chart_layout(fig_zu, "Participación de Utilidad por Zona", height=340)
        fig_zu.update_layout(legend=dict(orientation="v", font=dict(size=10)))
        st.plotly_chart(fig_zu, use_container_width=True)


# ══════════════════════════════════════════════
#  TAB 4 — EJECUTIVOS
# ══════════════════════════════════════════════
with tab4:
    st.markdown('<div class="section-header">Productividad por Ejecutivo</div>', unsafe_allow_html=True)

    ejec_gb = dff.groupby("EJECUTVO").agg(
        Venta=("TOTAL VENTA","sum"),
        Utilidad=("UTILIDAD TOTAL $","sum"),
        Descuento=("TOTAL DESCUENTO","sum"),
        Pares=("PARES","sum"),
        Facturas=("FACTURA_KEY","nunique"),
    ).reset_index()
    ejec_gb["Margen%"] = ejec_gb["Utilidad"] / ejec_gb["Venta"].replace(0,np.nan) * 100
    ejec_gb["Ticket"]  = ejec_gb["Venta"] / ejec_gb["Facturas"].replace(0, np.nan)
    ejec_gb["Desc%"]   = ejec_gb["Descuento"] / ejec_gb["Venta"].replace(0,np.nan) * 100
    ejec_gb = ejec_gb.sort_values("Venta", ascending=False)
    ejec_gb["Ejecutivo"] = ejec_gb["EJECUTVO"].str.title()

    col1, col2 = st.columns(2)
    with col1:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=ejec_gb["Ejecutivo"], y=ejec_gb["Venta"],
            name="Venta Total",
            marker_color=COLORS["gray4"],
            text=[fmt_usd(v) for v in ejec_gb["Venta"]],
            textposition="outside", textfont=dict(size=9),
        ))
        fig.add_trace(go.Bar(
            x=ejec_gb["Ejecutivo"], y=ejec_gb["Utilidad"],
            name="Utilidad",
            marker_color=COLORS["gray2"],
            text=[fmt_usd(v) for v in ejec_gb["Utilidad"]],
            textposition="outside", textfont=dict(size=9),
        ))
        chart_layout(fig, "Venta y Utilidad por Ejecutivo", height=380)
        fig.update_layout(barmode="group",
                          xaxis_tickangle=-30,
                          legend=dict(orientation="h", y=-0.25))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Radar de desempeño
        cats = ["Venta", "Utilidad", "Pares", "Facturas", "Ticket"]
        ejec_norm = ejec_gb.copy()
        for c in cats:
            mx = ejec_norm[c].max()
            ejec_norm[c+"_n"] = ejec_norm[c] / mx if mx > 0 else 0

        fig_r = go.Figure()
        for i, row in ejec_norm.iterrows():
            vals = [row[c+"_n"] for c in cats]
            vals += vals[:1]
            color = GRAY_SCALE[i % len(GRAY_SCALE)]
            fig_r.add_trace(go.Scatterpolar(
                r=vals,
                theta=cats + cats[:1],
                fill="toself",
                name=row["Ejecutivo"],
                line=dict(color=color, width=1.5),
                fillcolor=hex_to_rgba(color, 0.12),
                opacity=0.85,
            ))
        chart_layout(fig_r, "Radar de Desempeño Multi-Métrica", height=380)
        fig_r.update_layout(
            polar=dict(
                bgcolor=COLORS["card"],
                radialaxis=dict(visible=True, range=[0,1],
                                gridcolor=COLORS["gray1"],
                                tickfont=dict(size=8)),
                angularaxis=dict(gridcolor=COLORS["gray1"]),
            ),
            legend=dict(font=dict(size=9)),
        )
        st.plotly_chart(fig_r, use_container_width=True)

    # Margen y descuento por ejecutivo
    col3, col4 = st.columns(2)
    with col3:
        fig3 = go.Figure()
        ejec_s = ejec_gb.sort_values("Margen%", ascending=True)
        fig3.add_trace(go.Bar(
            x=ejec_s["Margen%"],
            y=ejec_s["Ejecutivo"],
            orientation="h",
            marker_color=[COLORS["gray4"] if v > ejec_gb["Margen%"].mean()
                          else COLORS["gray2"] for v in ejec_s["Margen%"]],
            text=[f"{v:.1f}%" for v in ejec_s["Margen%"]],
            textposition="outside",
        ))
        chart_layout(fig3, "Margen % por Ejecutivo (gris oscuro = sobre media)", height=340)
        fig3.update_layout(showlegend=False, xaxis_ticksuffix="%")
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        fig4 = go.Figure()
        ejec_d = ejec_gb.sort_values("Ticket", ascending=True)
        fig4.add_trace(go.Bar(
            x=ejec_d["Ticket"],
            y=ejec_d["Ejecutivo"],
            orientation="h",
            marker_color=COLORS["gray3"],
            text=[fmt_usd(v) for v in ejec_d["Ticket"]],
            textposition="outside",
        ))
        chart_layout(fig4, "Ticket Promedio por Ejecutivo", height=340)
        fig4.update_layout(showlegend=False)
        st.plotly_chart(fig4, use_container_width=True)

    # Tabla completa
    st.markdown('<div class="section-header">Tabla Resumen de Ejecutivos</div>', unsafe_allow_html=True)
    tbl = ejec_gb[["Ejecutivo","Venta","Utilidad","Descuento","Pares","Facturas","Margen%","Ticket","Desc%"]].copy()
    tbl["Venta"]    = tbl["Venta"].apply(fmt_usd)
    tbl["Utilidad"] = tbl["Utilidad"].apply(fmt_usd)
    tbl["Descuento"]= tbl["Descuento"].apply(fmt_usd)
    tbl["Ticket"]   = tbl["Ticket"].apply(fmt_usd)
    tbl["Margen%"]  = tbl["Margen%"].apply(lambda v: f"{v:.1f}%")
    tbl["Desc%"]    = tbl["Desc%"].apply(lambda v: f"{v:.1f}%")
    tbl.columns = ["Ejecutivo","Venta","Utilidad","Descuento","Unidades","Facturas",
                   "Margen %","Ticket Prom.","% Desc."]
    st.dataframe(tbl, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
#  TAB 5 — MARCAS & PRODUCTOS
# ══════════════════════════════════════════════
with tab5:
    st.markdown('<div class="section-header">Desempeño por Marca y Tipo de Producto</div>',
                unsafe_allow_html=True)

    marca_gb = dff.groupby("MARCA").agg(
        Venta=("TOTAL VENTA","sum"),
        Utilidad=("UTILIDAD TOTAL $","sum"),
        Descuento=("TOTAL DESCUENTO","sum"),
        Pares=("PARES","sum"),
    ).reset_index().sort_values("Venta", ascending=False)
    marca_gb["Margen%"] = marca_gb["Utilidad"] / marca_gb["Venta"].replace(0,np.nan) * 100
    marca_gb["Marca_L"] = marca_gb["MARCA"].str.title()

    top_n = st.slider("Top N marcas a mostrar", 5, len(marca_gb), min(15, len(marca_gb)))
    marca_top = marca_gb.head(top_n)

    col1, col2 = st.columns([3, 2])
    with col1:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=marca_top["Marca_L"], y=marca_top["Venta"],
            name="Venta", marker_color=COLORS["gray4"],
            text=[fmt_usd(v) for v in marca_top["Venta"]],
            textposition="outside", textfont=dict(size=9),
        ))
        fig.add_trace(go.Bar(
            x=marca_top["Marca_L"], y=marca_top["Utilidad"],
            name="Utilidad", marker_color=COLORS["gray2"],
            text=[fmt_usd(v) for v in marca_top["Utilidad"]],
            textposition="outside", textfont=dict(size=9),
        ))
        chart_layout(fig, f"Top {top_n} Marcas por Venta y Utilidad", height=380)
        fig.update_layout(barmode="group", xaxis_tickangle=-35,
                          legend=dict(orientation="h", y=-0.25))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = go.Figure(go.Pie(
            labels=marca_top["Marca_L"], values=marca_top["Venta"],
            hole=0.5,
            marker=dict(colors=[GRAY_SCALE[i % len(GRAY_SCALE)] for i in range(len(marca_top))],
                        line=dict(color="#FFF", width=1.5)),
            textfont=dict(size=9),
        ))
        chart_layout(fig2, "Participación de Mercado (Venta)", height=380)
        fig2.update_layout(legend=dict(font=dict(size=9)))
        st.plotly_chart(fig2, use_container_width=True)

    # Margen por marca (scatter)
    col3, col4 = st.columns(2)
    with col3:
        fig3 = go.Figure()
        for i, row in marca_top.iterrows():
            fig3.add_trace(go.Scatter(
                x=[row["Venta"]], y=[row["Margen%"]],
                mode="markers+text",
                marker=dict(size=max(8, row["Pares"] / marca_top["Pares"].max() * 40 + 6),
                            color=GRAY_SCALE[list(marca_top.index).index(i) % len(GRAY_SCALE)],
                            line=dict(color="#000", width=0.8)),
                text=[row["Marca_L"]],
                textposition="top center",
                textfont=dict(size=8),
                name=row["Marca_L"],
            ))
        chart_layout(fig3, "Venta vs Margen % por Marca (tamaño = unidades)", height=380)
        fig3.update_layout(showlegend=False,
                           xaxis_tickprefix="$", yaxis_ticksuffix="%")
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        # Por tipo de producto
        prod_gb = dff.groupby("TIPO DE PRODUCTO").agg(
            Venta=("TOTAL VENTA","sum"),
            Utilidad=("UTILIDAD TOTAL $","sum"),
        ).reset_index().sort_values("Venta", ascending=True)
        prod_gb["Margen%"] = prod_gb["Utilidad"] / prod_gb["Venta"].replace(0,np.nan) * 100
        prod_gb["Prod_L"] = prod_gb["TIPO DE PRODUCTO"].str.title()

        fig4 = go.Figure(go.Bar(
            x=prod_gb["Venta"],
            y=prod_gb["Prod_L"],
            orientation="h",
            marker_color=[GRAY_SCALE[i % len(GRAY_SCALE)] for i in range(len(prod_gb))],
            text=[fmt_usd(v) for v in prod_gb["Venta"]],
            textposition="outside",
        ))
        chart_layout(fig4, "Venta por Tipo de Producto", height=380)
        fig4.update_layout(showlegend=False)
        st.plotly_chart(fig4, use_container_width=True)

    # Tipo de venta
    st.markdown('<div class="section-header">Tipo de Venta: Contado vs Crédito</div>',
                unsafe_allow_html=True)
    tv_gb = dff.groupby(["TIPO DE VENTA","MES"]).agg(
        Venta=("TOTAL VENTA","sum"),
    ).reset_index()
    tv_gb["MES_NUM"] = tv_gb["MES"].map(
        {"ENERO":1,"FEBRERO":2,"MARZO":3,"ABRIL":4,"MAYO":5,"JUNIO":6,
         "JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,"OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12})
    tv_gb = tv_gb.sort_values("MES_NUM")
    tv_gb["MES_L"] = tv_gb["MES"].str.capitalize()

    fig5 = go.Figure()
    for i, tv in enumerate(tv_gb["TIPO DE VENTA"].unique()):
        s = tv_gb[tv_gb["TIPO DE VENTA"] == tv]
        fig5.add_trace(go.Scatter(
            x=s["MES_L"], y=s["Venta"],
            name=tv.title(),
            mode="lines+markers",
            line=dict(color=GRAY_SCALE[i * 2], width=2.5),
            fill="tozeroy",
            fillcolor=hex_to_rgba(GRAY_SCALE[i*2], 0.13),
        ))
    chart_layout(fig5, "Evolución: Contado vs Crédito por Mes", height=300)
    fig5.update_layout(legend=dict(orientation="h", y=-0.2))
    st.plotly_chart(fig5, use_container_width=True)


# ══════════════════════════════════════════════
#  TAB 6 — ANÁLISIS DE VARIACIONES (¿POR QUÉ?)
# ══════════════════════════════════════════════
with tab6:
    st.markdown('<div class="section-header">Análisis de Variaciones — ¿Por Qué Pasó?</div>',
                unsafe_allow_html=True)

    # Comparativa mes a mes
    mes_var = dff.groupby("MES").agg(
        Venta=("TOTAL VENTA","sum"),
        Utilidad=("UTILIDAD TOTAL $","sum"),
        Descuento=("TOTAL DESCUENTO","sum"),
        Costo=("COSTO_TOTAL","sum"),
        Pares=("PARES","sum"),
    ).reset_index()
    mes_var["MES_NUM"] = mes_var["MES"].map(
        {"ENERO":1,"FEBRERO":2,"MARZO":3,"ABRIL":4,"MAYO":5,"JUNIO":6,
         "JULIO":7,"AGOSTO":8,"SEPTIEMBRE":9,"OCTUBRE":10,"NOVIEMBRE":11,"DICIEMBRE":12})
    mes_var = mes_var.sort_values("MES_NUM")
    mes_var["MES_L"] = mes_var["MES"].str.capitalize()
    mes_var["Margen%"] = mes_var["Utilidad"] / mes_var["Venta"].replace(0,np.nan) * 100

    # Variación vs mes anterior
    for col in ["Venta","Utilidad","Descuento","Pares","Margen%"]:
        mes_var[f"Δ{col}"] = mes_var[col].pct_change() * 100

    # Heatmap de correlaciones
    num_cols = ["TOTAL VENTA","UTILIDAD TOTAL $","TOTAL DESCUENTO",
                "COSTO_TOTAL","PARES","MARGEN_%"]
    corr = dff[num_cols].corr()
    labels = ["Venta","Utilidad","Descuento","Costo","Pares","Margen%"]

    col1, col2 = st.columns(2)
    with col1:
        fig_heat = go.Figure(go.Heatmap(
            z=corr.values,
            x=labels, y=labels,
            colorscale=[
                [0,   "#FFFFFF"],
                [0.5, "#9A9A9A"],
                [1,   "#1A1A1A"],
            ],
            text=np.round(corr.values, 2),
            texttemplate="%{text}",
            textfont=dict(size=11, color="#111"),
            zmin=-1, zmax=1,
            showscale=True,
            colorbar=dict(title="Corr.", tickfont=dict(size=9)),
        ))
        chart_layout(fig_heat, "Mapa de Correlaciones entre Variables Clave", height=380)
        fig_heat.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(fig_heat, use_container_width=True)

    with col2:
        # Variación MoM de margen%
        mes_var2 = mes_var.dropna(subset=["ΔMargen%"])
        fig_var = go.Figure()
        fig_var.add_trace(go.Bar(
            x=mes_var2["MES_L"],
            y=mes_var2["ΔMargen%"],
            marker_color=[COLORS["gray5"] if v >= 0 else COLORS["gray2"]
                          for v in mes_var2["ΔMargen%"]],
            text=[f"{v:+.1f}%" for v in mes_var2["ΔMargen%"]],
            textposition="outside",
        ))
        chart_layout(fig_var, "Variación Mes a Mes del Margen % (MoM)", height=380)
        fig_var.update_layout(showlegend=False, yaxis_ticksuffix="%")
        st.plotly_chart(fig_var, use_container_width=True)

    # Impacto del descuento sobre la utilidad
    st.markdown('<div class="section-header">Impacto de Descuentos sobre Rentabilidad</div>',
                unsafe_allow_html=True)

    # Agrupar por % de descuento en rangos
    dff2 = dff.copy()
    dff2["% DE DESCUENTO"] = dff2["% DE DESCUENTO"].fillna(0)
    bins = [-0.01, 0.001, 0.05, 0.10, 0.20, 0.30, 1.01]
    labs = ["Sin dcto.", "1-5%", "5-10%", "10-20%", "20-30%", ">30%"]
    dff2["Rango_Dcto"] = pd.cut(dff2["% DE DESCUENTO"], bins=bins, labels=labs)

    dcto_rng = dff2.groupby("Rango_Dcto", observed=True).agg(
        Venta=("TOTAL VENTA","sum"),
        Utilidad=("UTILIDAD TOTAL $","sum"),
        Cnt=("FACTURA_KEY","count"),
    ).reset_index()
    dcto_rng["Margen%"] = dcto_rng["Utilidad"] / dcto_rng["Venta"].replace(0,np.nan) * 100

    col3, col4 = st.columns(2)
    with col3:
        fig6 = make_subplots(specs=[[{"secondary_y": True}]])
        fig6.add_trace(go.Bar(
            x=dcto_rng["Rango_Dcto"].astype(str),
            y=dcto_rng["Venta"],
            name="Venta", marker_color=COLORS["gray3"],
        ), secondary_y=False)
        fig6.add_trace(go.Scatter(
            x=dcto_rng["Rango_Dcto"].astype(str),
            y=dcto_rng["Margen%"],
            name="Margen%", mode="lines+markers+text",
            line=dict(color=COLORS["gray5"], width=2.5),
            marker=dict(size=8),
            text=[f"{v:.1f}%" for v in dcto_rng["Margen%"]],
            textposition="top center", textfont=dict(size=10),
        ), secondary_y=True)
        fig6.update_yaxes(title_text="Venta $", secondary_y=False, tickprefix="$")
        fig6.update_yaxes(title_text="Margen %", secondary_y=True, ticksuffix="%", showgrid=False)
        chart_layout(fig6, "Rango de Descuento vs Margen%: ¿a qué punto el descuento destruye margen?",
                     height=380)
        fig6.update_layout(legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig6, use_container_width=True)

    with col4:
        # Desglose de utilidad: por segmento y zona
        seg_zona = dff.groupby(["ZONA","SEGMENTO"]).agg(
            Utilidad=("UTILIDAD TOTAL $","sum"),
        ).reset_index()
        zonas = seg_zona["ZONA"].unique()

        fig7 = go.Figure()
        for i, seg in enumerate(seg_zona["SEGMENTO"].unique()):
            s = seg_zona[seg_zona["SEGMENTO"] == seg]
            fig7.add_trace(go.Bar(
                x=[z.title() for z in s["ZONA"]],
                y=s["Utilidad"],
                name=seg.title(),
                marker_color=GRAY_SCALE[i * 2 % len(GRAY_SCALE)],
            ))
        chart_layout(fig7, "Utilidad por Zona y Segmento: distribución de rentabilidad", height=380)
        fig7.update_layout(barmode="stack",
                           xaxis_tickangle=-25,
                           legend=dict(orientation="h", y=-0.25))
        st.plotly_chart(fig7, use_container_width=True)

    # Tabla variaciones MoM
    st.markdown('<div class="section-header">Variaciones Mes a Mes (MoM)</div>',
                unsafe_allow_html=True)
    tbl2 = mes_var[["MES_L","Venta","Utilidad","Descuento","Margen%",
                    "ΔVenta","ΔUtilidad","ΔDescuento","ΔMargen%"]].copy()
    for c in ["Venta","Utilidad","Descuento"]:
        tbl2[c] = tbl2[c].apply(fmt_usd)
    tbl2["Margen%"] = tbl2["Margen%"].apply(lambda v: f"{v:.1f}%")
    for c in ["ΔVenta","ΔUtilidad","ΔDescuento","ΔMargen%"]:
        tbl2[c] = tbl2[c].apply(lambda v: f"{v:+.1f}%" if pd.notna(v) else "—")
    tbl2.columns = ["Mes","Venta","Utilidad","Descuento","Margen%",
                    "Δ Venta","Δ Utilidad","Δ Descuento","Δ Margen%"]
    st.dataframe(tbl2, use_container_width=True, hide_index=True)

    # Insight automático
    st.markdown('<div class="section-header">Insights Automáticos</div>', unsafe_allow_html=True)

    mejor_zona  = dff.groupby("ZONA")["UTILIDAD TOTAL $"].sum().idxmax()
    peor_zona   = dff.groupby("ZONA")["UTILIDAD TOTAL $"].sum().idxmin()
    mejor_marca = dff.groupby("MARCA")["UTILIDAD TOTAL $"].sum().idxmax()
    mejor_ejec  = dff.groupby("EJECUTVO")["TOTAL VENTA"].sum().idxmax()
    mayor_desc_ejec = dff.groupby("EJECUTVO")["TOTAL DESCUENTO"].sum().idxmax()

    insights = [
        f"🏆 La zona con mayor utilidad acumulada es <b>{mejor_zona.title()}</b>, "
        f"mientras que <b>{peor_zona.title()}</b> es la de menor contribución.",
        f"📦 La marca más rentable del período es <b>{mejor_marca.title()}</b>.",
        f"👤 El ejecutivo con mayor volumen de venta es <b>{mejor_ejec.title()}</b>.",
        f"⚠️ <b>{mayor_desc_ejec.title()}</b> es el ejecutivo con mayor monto total "
        f"en descuentos otorgados — revisar política de descuentos.",
        f"📉 A mayor rango de descuento, el margen tiende a disminuir "
        f"(ver gráfico de Rango de Descuento vs Margen%).",
        f"📊 La correlación entre Descuento y Utilidad indica qué tanto afectan "
        f"los descuentos a la rentabilidad real del negocio.",
    ]
    for ins in insights:
        st.markdown(
            f'<div style="background:{COLORS["card"]};border:1px solid {COLORS["border"]};'
            f'border-left:3px solid {COLORS["gray5"]};border-radius:6px;'
            f'padding:12px 16px;margin-bottom:8px;font-size:0.85rem;'
            f'color:{COLORS["text_dark"]};">{ins}</div>',
            unsafe_allow_html=True,
        )

# ─────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────
st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)
st.markdown(f"""
<div style='border-top:1px solid {COLORS["border"]};padding:16px 0;
     display:flex;justify-content:space-between;
     font-size:0.70rem;color:{COLORS["text_light"]};'>
  <span>Dashboard Ejecutivo de Ventas 2026 · Diseño Gris Elegante</span>
  <span>Datos: Enero–Mayo 2026 · Actualización mensual automática</span>
</div>
""", unsafe_allow_html=True)
