"""
Dashboard editorial — Queimaduras no Brasil, 2025.
Para embed em matéria jornalística. Stack: Streamlit + Pandas + Plotly.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import os

# ═════════════════════════════════════════════════════════════════════════════
# 1. CONFIGURAÇÃO
# ═════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Queimaduras no Brasil • 2025",
    page_icon="🔥",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ═════════════════════════════════════════════════════════════════════════════
# 2. DESIGN SYSTEM — tokens
# ═════════════════════════════════════════════════════════════════════════════
INK         = "#1a1a1a"
INK_MUTED   = "#5a5a5a"
INK_LIGHT   = "#8a8a8a"
PAPER       = "#faf7f2"
CARD        = "#ffffff"
BORDER      = "#e8e2d6"
ACCENT      = "#b8492c"   # terracota — fogo/queimadura, sóbrio
ACCENT_DARK = "#7a2d18"
FEM         = "#c4806b"   # rosa-terroso, derivado do acento
MASC        = "#3d4f5e"   # azul-ardósia profundo

# Escala sequencial: paper → terracota → vermelho profundo
SCALE_TERRA = [
    [0.00, "#fbeee7"],
    [0.20, "#f0c9b6"],
    [0.45, "#dc8e6e"],
    [0.70, "#b8492c"],
    [1.00, "#6b2818"],
]

# Layout base para todos os gráficos Plotly — garante consistência visual
PLOTLY_LAYOUT = dict(
    font=dict(family="Inter, -apple-system, sans-serif", color=INK, size=13),
    paper_bgcolor=PAPER,
    plot_bgcolor=PAPER,
    title_font=dict(family="'Source Serif 4', Georgia, serif", color=INK, size=18),
    margin=dict(l=10, r=10, t=70, b=40),
    separators=",.",  # formato brasileiro: 1.234,56
    hoverlabel=dict(
        bgcolor=CARD,
        bordercolor=BORDER,
        font=dict(family="Inter, sans-serif", color=INK, size=12),
    ),
)

# ═════════════════════════════════════════════════════════════════════════════
# 3. CSS EDITORIAL
# ═════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,400;8..60,600;8..60,700&family=Inter:wght@400;500;600;700&display=swap');

/* ---- Base ---- */
html, body, [class*="css"] {{
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    color: {INK};
}}
.stApp {{ background: {PAPER}; }}
.block-container {{
    max-width: 1080px;
    padding-top: 2.5rem;
    padding-bottom: 3rem;
}}
#MainMenu, footer, header {{ visibility: hidden; height: 0; }}

/* ---- Tipografia editorial ---- */
.kicker {{
    font-family: 'Inter', sans-serif;
    text-transform: uppercase;
    letter-spacing: 0.20em;
    font-size: 0.76rem;
    font-weight: 700;
    color: {ACCENT};
    margin-bottom: 0.9rem;
}}
.headline {{
    font-family: 'Source Serif 4', Georgia, serif;
    font-size: 2.9rem;
    font-weight: 700;
    line-height: 1.08;
    letter-spacing: -0.018em;
    margin: 0 0 1.1rem 0;
    color: {INK};
}}
.dek {{
    font-family: 'Source Serif 4', Georgia, serif;
    font-size: 1.20rem;
    line-height: 1.55;
    color: {INK_MUTED};
    font-style: italic;
    margin-bottom: 1.4rem;
    max-width: 760px;
}}
.byline {{
    font-family: 'Inter', sans-serif;
    font-size: 0.82rem;
    color: {INK_MUTED};
    border-top: 1px solid {BORDER};
    border-bottom: 1px solid {BORDER};
    padding: 0.7rem 0;
    margin: 0.4rem 0 2.5rem 0;
    display: flex;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 0.5rem;
}}

/* ---- Cabeçalhos de seção ---- */
.section-num {{
    font-family: 'Source Serif 4', Georgia, serif;
    font-size: 0.95rem;
    font-weight: 700;
    color: {ACCENT};
    margin: 3rem 0 0.4rem 0;
    letter-spacing: 0.05em;
}}
.section-h {{
    font-family: 'Source Serif 4', Georgia, serif;
    font-size: 1.85rem;
    font-weight: 700;
    line-height: 1.2;
    letter-spacing: -0.012em;
    margin: 0 0 0.6rem 0;
    color: {INK};
}}
.section-sub {{
    font-family: 'Inter', sans-serif;
    font-size: 1rem;
    line-height: 1.65;
    color: {INK_MUTED};
    margin-bottom: 1.4rem;
    max-width: 760px;
}}
.section-sub b {{ color: {INK}; font-weight: 600; }}

/* ---- LEDE: 4 KPIs como cards ---- */
.lede-card {{
    background: {CARD};
    border: 1px solid {BORDER};
    border-radius: 2px;
    padding: 1.4rem 1.2rem 1.2rem 1.2rem;
    height: 100%;
    transition: border-color 0.2s ease;
}}
.lede-card:hover {{ border-color: {ACCENT}; }}
.lede-num {{
    font-family: 'Source Serif 4', serif;
    font-size: 2.5rem;
    font-weight: 700;
    line-height: 1;
    color: {ACCENT};
    margin: 0;
}}
.lede-label {{
    font-family: 'Inter', sans-serif;
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.14em;
    color: {INK_MUTED};
    margin: 0.6rem 0 0.7rem 0;
}}
.lede-context {{
    font-family: 'Source Serif 4', serif;
    font-size: 0.93rem;
    line-height: 1.5;
    color: {INK};
    margin: 0;
}}
.lede-context b {{ color: {ACCENT_DARK}; }}

/* ---- Pull-quote / callout ---- */
.callout {{
    border-left: 3px solid {ACCENT};
    padding: 0.5rem 0 0.5rem 1.3rem;
    margin: 1.6rem 0 0.6rem 0;
    font-family: 'Source Serif 4', serif;
    font-style: italic;
    color: {INK};
    font-size: 1.10rem;
    line-height: 1.55;
}}
.callout b {{ font-style: normal; color: {ACCENT_DARK}; }}

/* ---- Caption sob gráficos ---- */
.caption {{
    font-family: 'Inter', sans-serif;
    font-size: 0.78rem;
    color: {INK_LIGHT};
    margin: 0.2rem 0 2.4rem 0;
    line-height: 1.5;
}}

/* ---- Rodapé ---- */
.footer {{
    font-family: 'Inter', sans-serif;
    font-size: 0.78rem;
    color: {INK_MUTED};
    line-height: 1.65;
    border-top: 1px solid {BORDER};
    padding-top: 1.2rem;
    margin-top: 3rem;
    text-align: center;
}}

/* ---- Streamlit overrides ---- */
hr {{ border-color: {BORDER} !important; }}
.stRadio > label,
.stSelectbox label,
.stMultiselect label,
.stCheckbox label {{
    font-family: 'Inter', sans-serif !important;
    font-size: 0.84rem !important;
    color: {INK_MUTED} !important;
    font-weight: 500 !important;
}}
.stRadio [role="radiogroup"] {{ gap: 0.4rem; }}
.stRadio [role="radiogroup"] label {{
    background: {CARD};
    border: 1px solid {BORDER};
    border-radius: 2px;
    padding: 0.35rem 0.85rem;
    transition: all 0.15s ease;
}}
.stRadio [role="radiogroup"] label:hover {{
    border-color: {ACCENT};
    color: {ACCENT};
}}

/* Expander estilizado */
div[data-testid="stExpander"] {{
    border: 1px solid {BORDER};
    border-radius: 2px;
    background: {CARD};
    margin-bottom: 0.8rem;
}}
div[data-testid="stExpander"] summary {{
    font-family: 'Inter', sans-serif;
    font-size: 0.88rem;
    font-weight: 600;
    color: {INK};
}}
.stDataFrame {{ font-family: 'Inter', sans-serif; font-size: 0.85rem; }}
</style>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# 4. MAPEAMENTOS UF
# ═════════════════════════════════════════════════════════════════════════════
UF_CTQ = {
    'ACRE': 'AC', 'ALAGOAS': 'AL', 'AMAPÁ': 'AP', 'AMAZONAS': 'AM',
    'BAHIA': 'BA', 'CEARÁ': 'CE', 'DISTRITO FEDERAL': 'DF',
    'ESPÍRITO SANTO': 'ES', 'GOIÁS': 'GO', 'MARANHÃO': 'MA',
    'MATO GROSSO': 'MT', 'MATO GROSSO DO SUL': 'MS', 'MINAS GERAIS': 'MG',
    'PARANÁ': 'PR', 'PARAÍBA': 'PB', 'PARÁ': 'PA', 'PERNAMBUCO': 'PE',
    'PIAUÍ': 'PI', 'RIO GRANDE DO NORTE': 'RN', 'RIO GRANDE DO SUL': 'RS',
    'RIO DE JANEIRO': 'RJ', 'RONDÔNIA': 'RO', 'RORAIMA': 'RR',
    'SANTA CATARINA': 'SC', 'SERGIPE': 'SE', 'SÃO PAULO': 'SP',
    'TOCANTINS': 'TO',
}
UF_GRAF = {
    'ACRE': 'AC', 'ALAGOA': 'AL', 'AMAPÁ': 'AP', 'AMAZONA': 'AM',
    'BAHIA': 'BA', 'CEARÁ': 'CE', 'D. FEDERAL': 'DF', 'E. SANTO': 'ES',
    'GOIÁS': 'GO', 'MARANHÃO': 'MA', 'M. GROSSO': 'MT', 'M. G. SUL': 'MS',
    'M. GERAIS': 'MG', 'PARANÁ': 'PR', 'PARAIBA': 'PB', 'PARÁ': 'PA',
    'PERNAMBUCO': 'PE', 'PIAUÍ': 'PI', 'R. G. NORTE': 'RN',
    'R. G. SUL': 'RS', 'R. DE JANEIRO': 'RJ', 'RONDONIA': 'RO',
    'RORAIMA': 'RR', 'S. CATARINA': 'SC', 'SERGIPE': 'SE',
    'S. PAULO': 'SP', 'TOCANTINS': 'TO',
}
UF_NOME = {
    'AC': 'Acre', 'AL': 'Alagoas', 'AP': 'Amapá', 'AM': 'Amazonas',
    'BA': 'Bahia', 'CE': 'Ceará', 'DF': 'Distrito Federal',
    'ES': 'Espírito Santo', 'GO': 'Goiás', 'MA': 'Maranhão',
    'MT': 'Mato Grosso', 'MS': 'Mato Grosso do Sul', 'MG': 'Minas Gerais',
    'PR': 'Paraná', 'PB': 'Paraíba', 'PA': 'Pará', 'PE': 'Pernambuco',
    'PI': 'Piauí', 'RN': 'Rio Grande do Norte', 'RS': 'Rio Grande do Sul',
    'RJ': 'Rio de Janeiro', 'RO': 'Rondônia', 'RR': 'Roraima',
    'SC': 'Santa Catarina', 'SE': 'Sergipe', 'SP': 'São Paulo',
    'TO': 'Tocantins',
}

# ═════════════════════════════════════════════════════════════════════════════
# 5. LOADERS — cacheados, robustos a variantes de nome de arquivo
# ═════════════════════════════════════════════════════════════════════════════
def _find(*candidates):
    """Devolve o primeiro nome de arquivo existente entre as variantes."""
    for c in candidates:
        p = os.path.join(BASE_DIR, c)
        if os.path.exists(p):
            return p
    raise FileNotFoundError(f"Arquivo não encontrado: {candidates}")


@st.cache_data
def load_ctq():
    df = pd.read_excel(_find('ctq.xlsx'))
    df.columns = ['Estado', 'CTQ']
    df['UF']   = df['Estado'].map(UF_CTQ)
    df['Nome'] = df['UF'].map(UF_NOME)
    return df


@st.cache_data
def load_faixa():
    df = pd.read_excel(
        _find('Faixa Etaria (1).xlsx', 'Faixa_Etaria__1_.xlsx', 'Faixa Etaria.xlsx'),
        header=1,
    )
    df.columns = ['Faixa', 'Feminino', 'Masculino']
    df['Feminino']  = pd.to_numeric(df['Feminino'],  errors='coerce')
    df['Masculino'] = pd.to_numeric(df['Masculino'], errors='coerce')
    df = df.dropna(subset=['Faixa', 'Feminino', 'Masculino']).reset_index(drop=True)
    df['Total'] = df['Feminino'] + df['Masculino']
    return df


@st.cache_data
def load_intern():
    df = pd.read_excel(_find('GRAFICO2.xlsx'))
    df.columns = ['Estado', 'Internacao', 'Morte']
    df['Morte']       = pd.to_numeric(df['Morte'], errors='coerce').fillna(0).astype(int)
    df['UF']          = df['Estado'].map(UF_GRAF)
    df['Nome']        = df['UF'].map(UF_NOME)
    df['Mortalidade'] = (df['Morte'] / df['Internacao'] * 100).round(2)
    return df


@st.cache_data
def load_geo():
    with open(_find('brazil_states.geojson'), 'r', encoding='utf-8') as f:
        return json.load(f)


df_ctq  = load_ctq()
df_fe   = load_faixa()
df_int  = load_intern()
geojson = load_geo()

# Merge para o mapa
df_mapa = df_int.merge(df_ctq[['UF', 'CTQ']], on='UF', how='left')

# Agregados
TOT_INT  = int(df_int['Internacao'].sum())
TOT_OB   = int(df_int['Morte'].sum())
TOT_CTQ  = int(df_ctq['CTQ'].sum())
TAX_MORT = round(TOT_OB / TOT_INT * 100, 2)
SEM_CTQ  = df_ctq.loc[df_ctq['CTQ'] == 0, 'Nome'].tolist()
SP_CTQ   = int(df_ctq.loc[df_ctq['UF'] == 'SP', 'CTQ'].values[0])
UFS_COM_CTQ = int((df_ctq['CTQ'] > 0).sum())

# Helpers
def br(n):
    """Formata número estilo brasileiro: 1.234.567"""
    return f"{int(n):,}".replace(",", ".")


def pretty_faixa(s, short=False):
    """
    Normaliza códigos do DATASUS para leitura jornalística.
    `short=True` devolve forma curta para eixos de gráficos.
        '<1a'    → 'Menos de 1 ano'   ou  '<1'
        '1-4a'   → '1-4 anos'         ou  '1-4'
        '80e+a'  → '80+ anos'         ou  '80+'
    """
    if s == '<1a':
        return '<1' if short else 'Menos de 1 ano'
    if s.endswith('e+a'):
        base = s.replace('e+a', '+')
        return base if short else f"{base} anos"
    if s.endswith('a'):
        base = s[:-1]
        return base if short else f"{base} anos"
    return s


# ═════════════════════════════════════════════════════════════════════════════
# 6. CABEÇALHO EDITORIAL
# ═════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="kicker">Saúde pública · Brasil · 2025</div>', unsafe_allow_html=True)
st.markdown(
    '<h1 class="headline">As queimaduras que o Brasil<br>não está vendo</h1>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<p class="dek">Em 2025, {br(TOT_INT)} brasileiros foram internados '
    f'por queimaduras. Quase mil não voltaram para casa. Por trás dos '
    f'números, um mapa desigual de risco, exposição e acesso a cuidado '
    f'especializado.</p>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="byline">'
    '<span>Dashboard interativo · Dados do Ministério da Saúde / DATASUS</span>'
    '<span>Recorte: ano de 2025</span>'
    '</div>',
    unsafe_allow_html=True,
)


# ═════════════════════════════════════════════════════════════════════════════
# 7. LEDE — quatro números em contexto
# ═════════════════════════════════════════════════════════════════════════════
faixa_pico_raw = df_fe.loc[df_fe['Total'].idxmax(), 'Faixa']
faixa_pico     = pretty_faixa(faixa_pico_raw)
total_pico     = int(df_fe['Total'].max())
pct_pico       = total_pico / TOT_INT * 100
minutos_int    = round(525_600 / TOT_INT)   # minutos em um ano / internações

c1, c2, c3, c4 = st.columns(4, gap="medium")
with c1:
    st.markdown(f"""
    <div class="lede-card">
        <p class="lede-num">{br(TOT_INT)}</p>
        <p class="lede-label">Internações</p>
        <p class="lede-context">Uma a cada <b>{minutos_int} minutos</b> no país.</p>
    </div>
    """, unsafe_allow_html=True)
with c2:
    st.markdown(f"""
    <div class="lede-card">
        <p class="lede-num">{br(TOT_OB)}</p>
        <p class="lede-label">Óbitos</p>
        <p class="lede-context">Mortalidade média de <b>{str(TAX_MORT).replace('.',',')}%</b> — varia até <b>5×</b> entre estados.</p>
    </div>
    """, unsafe_allow_html=True)
with c3:
    st.markdown(f"""
    <div class="lede-card">
        <p class="lede-num">{TOT_CTQ}</p>
        <p class="lede-label">Centros especializados</p>
        <p class="lede-context">Apenas <b>{UFS_COM_CTQ} dos 27 estados</b> contam com CTQ.</p>
    </div>
    """, unsafe_allow_html=True)
with c4:
    st.markdown(f"""
    <div class="lede-card">
        <p class="lede-num">{faixa_pico}</p>
        <p class="lede-label">Faixa mais atingida</p>
        <p class="lede-context">Crianças nessa idade são <b>{pct_pico:.0f}%</b> de todas as internações.</p>
    </div>
    """, unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# 8. SEÇÃO 1 — GEOGRAFIA (choropleth)
# ═════════════════════════════════════════════════════════════════════════════
top5 = df_int.nlargest(5, 'Internacao')
top5_total = int(top5['Internacao'].sum())
top5_pct = top5_total / TOT_INT * 100
top5_nomes = ', '.join(top5['Nome'].tolist())

st.markdown('<p class="section-num">1 · Geografia</p>', unsafe_allow_html=True)
st.markdown(
    '<h2 class="section-h">Cinco estados concentram mais da metade das internações</h2>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<p class="section-sub">{top5_nomes} somam <b>{br(top5_total)}</b> internações — '
    f'<b>{top5_pct:.0f}%</b> do total nacional. Escolha um indicador abaixo para '
    f'ver como ele se distribui pelo território.</p>',
    unsafe_allow_html=True,
)

variavel = st.radio(
    label="Indicador exibido no mapa",
    options=["Internações", "Óbitos", "Centros de Tratamento"],
    horizontal=True,
    label_visibility="collapsed",
)
col_var = {
    "Internações": "Internacao",
    "Óbitos": "Morte",
    "Centros de Tratamento": "CTQ",
}[variavel]

fig_map = go.Figure(go.Choropleth(
    geojson=geojson,
    locations=df_mapa['UF'],
    z=df_mapa[col_var],
    featureidkey='properties.sigla',
    colorscale=SCALE_TERRA,
    marker_line_color=PAPER,
    marker_line_width=0.9,
    colorbar=dict(
        title="",
        thickness=8,
        len=0.55,
        x=0.96,
        y=0.45,
        tickfont=dict(family="Inter", size=11, color=INK_MUTED),
        outlinewidth=0,
    ),
    customdata=df_mapa[['Nome', 'Internacao', 'Morte', 'CTQ', 'Mortalidade']].values,
    hovertemplate=(
        '<b>%{customdata[0]}</b><br>'
        '<span style="color:#5a5a5a">Internações:</span> %{customdata[1]:,}<br>'
        '<span style="color:#5a5a5a">Óbitos:</span> %{customdata[2]:,}<br>'
        '<span style="color:#5a5a5a">CTQs:</span> %{customdata[3]}<br>'
        '<span style="color:#5a5a5a">Mortalidade:</span> %{customdata[4]:.2f}%'
        '<extra></extra>'
    ),
))
fig_map.update_geos(fitbounds='locations', visible=False, bgcolor=PAPER)
fig_map.update_layout(
    **PLOTLY_LAYOUT,
    height=540,
    title=dict(
        text=f"<b>{variavel} por estado</b> &nbsp;&nbsp;<span style='font-size:13px;color:#8a8a8a;font-family:Inter'>passe o mouse para detalhes</span>",
        x=0.02, y=0.97, xanchor='left',
    ),
)
st.plotly_chart(fig_map, width='stretch', config={"displayModeBar": False})

st.markdown(
    '<p class="caption">Tonalidade mais escura indica valor mais alto. '
    'Fonte: Ministério da Saúde / DATASUS, 2025.</p>',
    unsafe_allow_html=True,
)


# ═════════════════════════════════════════════════════════════════════════════
# 9. SEÇÃO 2 — QUEM SE QUEIMA (pirâmide etária)
# ═════════════════════════════════════════════════════════════════════════════
pct_masc = df_fe['Masculino'].sum() / TOT_INT * 100
# Adultos: 20 anos ou mais (linhas 5+ do dataframe; exclui <1, 1-4, 5-9, 10-14, 15-19)
df_adultos = df_fe.iloc[5:]
faixa_adulto_max = pretty_faixa(df_adultos.loc[df_adultos['Total'].idxmax(), 'Faixa'])

st.markdown('<p class="section-num">2 · Quem se queima</p>', unsafe_allow_html=True)
st.markdown(
    '<h2 class="section-h">Crianças pequenas e homens em idade produtiva concentram o risco</h2>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<p class="section-sub">Crianças de <b>1 a 4 anos</b> são a faixa etária mais '
    f'atingida — sozinhas, representam <b>{pct_pico:.0f}%</b> de todas as internações. '
    f'Entre adultos, o pico está em <b>{faixa_adulto_max}</b>. No total, '
    f'<b>{pct_masc:.0f}% das vítimas são homens</b>.</p>',
    unsafe_allow_html=True,
)

# Rótulos curtos para a pirâmide ("<1", "1-4", ..., "80+")
df_fe_disp = df_fe.copy()
df_fe_disp['Faixa_Display'] = df_fe_disp['Faixa'].apply(lambda s: pretty_faixa(s, short=True))

# Pirâmide etária (barras espelhadas)
fig_pyr = go.Figure()
fig_pyr.add_trace(go.Bar(
    y=df_fe_disp['Faixa_Display'],
    x=-df_fe_disp['Masculino'],
    name='Masculino',
    orientation='h',
    marker=dict(color=MASC, line=dict(color=PAPER, width=0.5)),
    customdata=df_fe_disp['Masculino'],
    hovertemplate='<b>Masculino · %{y} anos</b><br>Internações: %{customdata:,}<extra></extra>',
))
fig_pyr.add_trace(go.Bar(
    y=df_fe_disp['Faixa_Display'],
    x=df_fe_disp['Feminino'],
    name='Feminino',
    orientation='h',
    marker=dict(color=FEM, line=dict(color=PAPER, width=0.5)),
    hovertemplate='<b>Feminino · %{y} anos</b><br>Internações: %{x:,}<extra></extra>',
))

max_x = max(df_fe['Masculino'].max(), df_fe['Feminino'].max())
tick_step = 1000
tick_vals = list(range(-int(max_x // tick_step + 1) * tick_step,
                       int(max_x // tick_step + 1) * tick_step + 1,
                       tick_step))
tick_text = [f"{abs(v):,}".replace(",", ".") if v != 0 else "0" for v in tick_vals]

fig_pyr.update_layout(
    **PLOTLY_LAYOUT,
    height=580,
    barmode='overlay',
    bargap=0.22,
    showlegend=False,
    title=dict(text="<b>Pirâmide etária das internações</b>", x=0.02, y=0.97, xanchor='left'),
    xaxis=dict(
        tickvals=tick_vals,
        ticktext=tick_text,
        tickfont=dict(family="Inter", size=11, color=INK_MUTED),
        showgrid=True, gridcolor=BORDER, gridwidth=0.5,
        zeroline=True, zerolinecolor=INK_MUTED, zerolinewidth=1,
        title="",
    ),
    yaxis=dict(
        tickfont=dict(family="Inter", size=12, color=INK),
        showgrid=False,
        autorange='reversed',  # menores idades no topo
        title="",
    ),
)
# Rótulos de lado (substitui legenda)
fig_pyr.add_annotation(
    text=f"<b style='color:{MASC}'>◀ MASCULINO</b>",
    xref="paper", yref="paper", x=0.18, y=1.04, showarrow=False,
    font=dict(family="Inter", size=11, color=MASC),
)
fig_pyr.add_annotation(
    text=f"<b style='color:{FEM}'>FEMININO ▶</b>",
    xref="paper", yref="paper", x=0.82, y=1.04, showarrow=False,
    font=dict(family="Inter", size=11, color=FEM),
)

st.plotly_chart(fig_pyr, width='stretch', config={"displayModeBar": False})

st.markdown(
    '<p class="caption">Barras espelhadas indicam volume absoluto de internações em cada faixa, '
    'com idades menores no topo e maiores na base. Fonte: Ministério da Saúde.</p>',
    unsafe_allow_html=True,
)


# ═════════════════════════════════════════════════════════════════════════════
# 10. SEÇÃO 3 — LETALIDADE (scatter)
# ═════════════════════════════════════════════════════════════════════════════
rj_mort = float(df_int.loc[df_int['UF'] == 'RJ', 'Mortalidade'].values[0])

st.markdown('<p class="section-num">3 · Letalidade</p>', unsafe_allow_html=True)
st.markdown(
    '<h2 class="section-h">No Rio, taxa de mortalidade quase dobra a média nacional</h2>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<p class="section-sub">A taxa de mortalidade hospitalar média no país é de '
    f'<b>{str(TAX_MORT).replace(".",",")}%</b>, mas o Rio de Janeiro registra '
    f'<b>{str(rj_mort).replace(".",",")}%</b> — o pior número entre estados com '
    f'volume significativo. No gráfico, cada ponto é um estado: à direita estão os '
    f'que mais internam, no alto os que perdem proporcionalmente mais vidas.</p>',
    unsafe_allow_html=True,
)

# Roraima excluída por dado de óbito ausente
df_scat = df_int[df_int['UF'] != 'RR'].copy()
# Rotula apenas os outliers (top em mortalidade ou volume)
df_scat['rot_y'] = df_scat['Mortalidade'].rank(ascending=False) <= 5
df_scat['rot_x'] = df_scat['Internacao'].rank(ascending=False) <= 3
df_scat['label'] = df_scat.apply(
    lambda r: f" {r['UF']}" if (r['rot_y'] or r['rot_x']) else "", axis=1,
)

fig_sc = go.Figure()
fig_sc.add_trace(go.Scatter(
    x=df_scat['Internacao'],
    y=df_scat['Mortalidade'],
    mode='markers+text',
    text=df_scat['label'],
    textposition='middle right',
    textfont=dict(family="Inter", size=11, color=INK, weight=600),
    marker=dict(
        size=14,
        color=df_scat['Mortalidade'],
        colorscale=SCALE_TERRA,
        cmin=0, cmax=df_scat['Mortalidade'].max(),
        line=dict(color=PAPER, width=1.2),
        showscale=False,
    ),
    customdata=df_scat[['Nome', 'Morte']].values,
    hovertemplate=(
        '<b>%{customdata[0]}</b><br>'
        '<span style="color:#5a5a5a">Internações:</span> %{x:,}<br>'
        '<span style="color:#5a5a5a">Óbitos:</span> %{customdata[1]:,}<br>'
        '<span style="color:#5a5a5a">Mortalidade:</span> %{y:.2f}%'
        '<extra></extra>'
    ),
))
# Linha da média nacional
fig_sc.add_hline(
    y=TAX_MORT,
    line_dash="dash", line_color=INK_MUTED, line_width=1,
    annotation_text=f"  Média nacional: {str(TAX_MORT).replace('.',',')}%",
    annotation_position="top left",
    annotation_font=dict(family="Inter", size=11, color=INK_MUTED),
)

fig_sc.update_layout(
    **PLOTLY_LAYOUT,
    height=500,
    title=dict(
        text="<b>Letalidade × volume de internações</b>",
        x=0.02, y=0.97, xanchor='left',
    ),
    xaxis=dict(
        title=dict(text="Internações em 2025", font=dict(family="Inter", size=12, color=INK_MUTED)),
        showgrid=True, gridcolor=BORDER, gridwidth=0.5,
        zeroline=False,
        tickformat=",d",
        tickfont=dict(family="Inter", size=11, color=INK_MUTED),
    ),
    yaxis=dict(
        title=dict(text="Taxa de mortalidade hospitalar", font=dict(family="Inter", size=12, color=INK_MUTED)),
        showgrid=True, gridcolor=BORDER, gridwidth=0.5,
        zeroline=False,
        ticksuffix="%",
        tickfont=dict(family="Inter", size=11, color=INK_MUTED),
    ),
)
st.plotly_chart(fig_sc, width='stretch', config={"displayModeBar": False})

st.markdown(
    '<p class="caption">Rótulos visíveis nos estados com maior mortalidade ou maior volume. '
    'Roraima excluída por ausência de dado de óbito no recorte público. '
    'Fonte: Ministério da Saúde.</p>',
    unsafe_allow_html=True,
)


# ═════════════════════════════════════════════════════════════════════════════
# 11. SEÇÃO 4 — INFRAESTRUTURA (CTQ)
# ═════════════════════════════════════════════════════════════════════════════
sp_share = SP_CTQ / TOT_CTQ * 100

st.markdown('<p class="section-num">4 · Infraestrutura</p>', unsafe_allow_html=True)
st.markdown(
    '<h2 class="section-h">Sete estados não têm sequer um centro especializado</h2>',
    unsafe_allow_html=True,
)
st.markdown(
    f'<p class="section-sub">O Brasil dispõe de <b>{TOT_CTQ} Centros de Tratamento de '
    f'Queimados</b> (CTQ) de média e alta gravidade. Mas <b>{SP_CTQ} deles '
    f'({sp_share:.0f}%) estão em São Paulo</b>. Outros sete estados, concentrados '
    f'no Norte e Nordeste, dependem de transferências para unidades em outras '
    f'unidades da federação.</p>',
    unsafe_allow_html=True,
)

df_ctq_sorted = df_ctq.sort_values('CTQ', ascending=True).reset_index(drop=True)

fig_ctq = go.Figure(go.Bar(
    x=df_ctq_sorted['CTQ'],
    y=df_ctq_sorted['Nome'],
    orientation='h',
    marker=dict(
        color=df_ctq_sorted['CTQ'],
        colorscale=SCALE_TERRA,
        line=dict(color=PAPER, width=0.5),
        showscale=False,
    ),
    text=df_ctq_sorted['CTQ'].apply(lambda v: f" {v}" if v > 0 else " 0"),
    textposition='outside',
    textfont=dict(family="Inter", size=11, color=INK),
    hovertemplate='<b>%{y}</b><br>CTQs: %{x}<extra></extra>',
))
fig_ctq.update_layout(
    **PLOTLY_LAYOUT,
    height=640,
    title=dict(
        text="<b>Centros de Tratamento de Queimados por estado</b>",
        x=0.02, y=0.98, xanchor='left',
    ),
    xaxis=dict(
        showgrid=True, gridcolor=BORDER, gridwidth=0.5,
        zeroline=True, zerolinecolor=BORDER,
        title="",
        tickfont=dict(family="Inter", size=11, color=INK_MUTED),
        range=[0, df_ctq['CTQ'].max() * 1.18],
    ),
    yaxis=dict(
        showgrid=False, zeroline=False, title="",
        tickfont=dict(family="Inter", size=11, color=INK),
    ),
)
st.plotly_chart(fig_ctq, width='stretch', config={"displayModeBar": False})

# Pull-quote com os 7 estados sem CTQ
st.markdown(
    f'<div class="callout">Sem qualquer centro especializado: <b>{", ".join(SEM_CTQ)}</b>.</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<p class="caption">Inclui hospitais públicos e filantrópicos contratualizados pelo SUS. '
    'Fonte: Ministério da Saúde.</p>',
    unsafe_allow_html=True,
)


# ═════════════════════════════════════════════════════════════════════════════
# 12. METODOLOGIA + TABELA COMPLETA (expanders)
# ═════════════════════════════════════════════════════════════════════════════
st.markdown('<p class="section-num">Sobre estes dados</p>', unsafe_allow_html=True)

with st.expander("Metodologia e fontes"):
    st.markdown(
        """
**Fonte primária.** Ministério da Saúde — Sistema de Informações Hospitalares
(SIH/SUS), DATASUS, exercício de 2025.

**Internações e óbitos.** Referem-se a hospitalizações por queimaduras
registradas no SUS no período. Roraima não registra óbitos no recorte
público; foi tratado como zero no cálculo agregado.

**Taxa de mortalidade.** Calculada como óbitos ÷ internações × 100 no mesmo
período. É uma medida de **letalidade hospitalar**, não de mortalidade
populacional.

**Centros de Tratamento de Queimados (CTQ).** Unidades hospitalares
habilitadas pelo Ministério da Saúde para atendimento de média e alta
gravidade. Inclui hospitais públicos e filantrópicos contratualizados pelo SUS.

**Faixa etária e sexo.** Consolidação nacional das internações distribuídas
em 18 faixas etárias e sexo registrado no atendimento.

**Limitações.** Os números refletem apenas hospitalizações registradas no SUS —
atendimentos ambulatoriais, atendimentos privados e queimaduras não
notificadas ficam fora da base. A subnotificação em estados com menor
capacidade de registro pode atenuar diferenças regionais reais.
        """
    )

with st.expander("Tabela completa por estado"):
    df_show = df_mapa[['Nome', 'Internacao', 'Morte', 'Mortalidade', 'CTQ']].copy()
    df_show.columns = ['Estado', 'Internações', 'Óbitos', 'Mortalidade (%)', 'CTQs']
    df_show = df_show.sort_values('Internações', ascending=False).reset_index(drop=True)
    st.dataframe(df_show, width='stretch', hide_index=True)


# ═════════════════════════════════════════════════════════════════════════════
# 13. RODAPÉ
# ═════════════════════════════════════════════════════════════════════════════
st.markdown(
    '<div class="footer">'
    'Dashboard interativo · Dados do Ministério da Saúde / DATASUS · '
    'Recorte: 2025 · Visualização desenvolvida em Streamlit + Plotly'
    '</div>',
    unsafe_allow_html=True,
)
