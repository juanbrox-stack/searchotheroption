"""
Gestor de Sustitutos Cecotec
- Sinstocks col K (índice 10) = SKU/Referencia Cecotec
- Tarifa Nacional T_AMZ: col REFERENCIA → cruce, col PVP PUB. = precio
- Tarifa Inter ES-FR/ES-IT etc: col REFERENCIA → cruce, col PVP PUB = precio
- Stock España: col Stock Operativo (col G)
- Stock FR/IT/DE/IT: col StockDisponible (col F)
- Sustituto: misma SUBFAMILIA, no desposicionado, PVP PUB ≤ original + 10€, stock > 0
- Salida: plantilla NUMBER / ARTICLE / NEW_ARTICLE
"""
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import re, io, glob, tempfile
from pathlib import Path

st.set_page_config(page_title="Sustitutos Cecotec", page_icon="🔄", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@600;700;800;900&family=Nunito+Sans:wght@400;600&display=swap');
:root{--blue:#3EB1C8;--black:#141413;--bg:#FAF9F5;--white:#fff;--grey:#e8e8e4;--muted:#6b7280;}
html,body,[class*="css"]{font-family:'Nunito Sans',sans-serif;background:var(--bg)!important;}
.stApp{background:var(--bg)!important;}
.navbar{background:var(--black);padding:.7rem 2rem;display:flex;align-items:center;gap:1.2rem;margin:-1rem -1rem 1.8rem -1rem;border-bottom:3px solid var(--blue);}
.navbar .logo{font-family:'Nunito',sans-serif;font-size:1.55rem;font-weight:900;color:#fff;letter-spacing:-.5px;}
.navbar .logo span{color:var(--blue);}
.navbar .sub{font-size:.72rem;color:rgba(255,255,255,.55);font-weight:600;text-transform:uppercase;letter-spacing:.12em;border-left:1px solid rgba(255,255,255,.2);padding-left:1.1rem;}
.sec{font-family:'Nunito',sans-serif;font-size:1rem;font-weight:800;color:var(--black);text-transform:uppercase;letter-spacing:.08em;padding:.4rem 0 .4rem .7rem;border-left:4px solid var(--blue);margin:1.4rem 0 .9rem;background:linear-gradient(90deg,rgba(62,177,200,.07),transparent);}
.kpi-row{display:flex;gap:.8rem;margin-bottom:1.4rem;flex-wrap:wrap;}
.kpi{background:var(--white);border-radius:10px;padding:.9rem 1.2rem;box-shadow:0 1px 3px rgba(20,20,19,.08),0 0 0 1px var(--grey);flex:1;min-width:110px;position:relative;overflow:hidden;}
.kpi::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;background:var(--blue);}
.kpi .val{font-size:1.9rem;font-weight:900;color:var(--blue);font-family:'Nunito',sans-serif;line-height:1;}
.kpi .lbl{font-size:.68rem;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin-top:.3rem;font-weight:600;}
div[data-testid="stTabs"] button[aria-selected="true"]{color:var(--blue)!important;border-bottom-color:var(--blue)!important;}
div[data-testid="stButton"] button[kind="primary"]{background:var(--blue)!important;border-color:var(--blue)!important;color:#fff!important;font-weight:800;border-radius:6px;}
div[data-testid="stDataFrame"] thead th{background:var(--black)!important;color:#fff!important;font-weight:700;font-size:.78rem;text-transform:uppercase;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="navbar"><div class="logo">ceco<span>tec</span></div><div class="sub">Gestor de sustitutos · Sin stock</div></div>', unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────────────────────
VALID_COUNTRIES = {'ES','FR','IT','DE','PT','PL','NL','BE'}
AMAZON_KEYWORDS = ['amazon','turaco-amazon','mano-a-mano']
# Tarifa inter: country → (sheet_name, pvp_col, neto_col)
INTER_SHEET = {
    'FR': ('ES-FR', 'PVP PUB', 'NETO ES-FR'),
    'IT': ('ES-IT', 'PVP PUB', 'NETO ES-IT'),
    'DE': ('ES-DE', 'PVP PUB', 'NETO ES-DE'),
    'PT': ('PT',    'PVP PUB', 'NETO PT'),
    'PL': ('PL',    'PVP PUB', 'NETO PL'),
    'NL': ('NL',    'PVP PUB', 'NETO NL'),
    'BE': ('BE',    'PVP PUB', 'NETO BE'),
}
UPLOAD_DIR = Path(__file__).parent

# ── Normalise reference ────────────────────────────────────────────────────────
def norm_ref(v) -> str:
    """Normalize reference: strip leading zeros for numeric, preserve A-prefix refs."""
    s = str(v).strip()
    if re.match(r'^A\d{2}_\w+_\d+$', s):
        return s  # A01_EU01_106744 — keep as-is
    if re.match(r'^\d+\.?\d*$', s):
        return str(int(float(s)))  # strip leading zeros
    return s

# ── Data loaders ───────────────────────────────────────────────────────────────
@st.cache_data
def load_tarifa_nac(path):
    df = pd.read_excel(path, sheet_name='T_AMZ')
    df['REF'] = df['REFERENCIA'].apply(norm_ref)
    return df

@st.cache_data
def load_tarifa_inter(path):
    xl = pd.ExcelFile(path)
    out = {}
    for sh in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sh)
        if 'REFERENCIA' not in df.columns:
            continue
        df['REF'] = df['REFERENCIA'].apply(norm_ref)
        out[sh] = df
    return out

@st.cache_data
def load_stock(path):
    """Returns dict country → DataFrame.
    España: col G = Stock Operativo
    FR/IT/DE/IT: col F = StockDisponible
    """
    xl = pd.ExcelFile(path)
    SHEET_MAP = {'España':'ES','Alemania':'DE','Francia':'FR','Italia':'IT'}
    out = {}
    for sh in xl.sheet_names:
        country = SHEET_MAP.get(sh, sh)
        df = pd.read_excel(path, sheet_name=sh)
        df['REF'] = df['Referencia'].apply(norm_ref)
        # Normalise stock column name
        if 'Stock Operativo' in df.columns:       # España
            df['STOCK'] = pd.to_numeric(df['Stock Operativo'], errors='coerce').fillna(0)
        elif 'StockDisponible' in df.columns:      # FR/IT/DE
            df['STOCK'] = pd.to_numeric(df['StockDisponible'], errors='coerce').fillna(0)
        else:
            df['STOCK'] = 0
        out[country] = df
    return out

@st.cache_data
def load_listing(path):
    """SKU → set of countries from Amazon orders file (sku + ship-country cols)."""
    p = Path(path)
    sku_countries = {}
    if p.suffix.lower() in ('.txt','.tsv','.csv'):
        df = pd.read_csv(path, sep='\t', encoding='utf-8-sig', dtype=str)
        cols = [c.strip() for c in df.columns]; df.columns = cols
        if 'sku' in cols and 'ship-country' in cols:
            for _, row in df[['sku','ship-country']].dropna().iterrows():
                country = str(row['ship-country']).strip().upper()
                if country not in VALID_COUNTRIES: continue
                m = re.search(r'(\d{3,6})$', str(row['sku']))
                if m:
                    sku_countries.setdefault(str(int(m.group(1))), set()).add(country)
        elif 'SKU del vendedor' in cols:
            for sku in df['SKU del vendedor'].dropna():
                s = str(sku).strip()
                m = re.match(r'^([A-Z]{2})0*(\d+)$', s)
                if m and m.group(1) in VALID_COUNTRIES:
                    sku_countries.setdefault(str(int(m.group(2))), set()).add(m.group(1))
                else:
                    m2 = re.search(r'(\d{3,6})$', s)
                    if m2: sku_countries.setdefault(str(int(m2.group(1))), set()).add('ES')
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb['BBDD'] if 'BBDD' in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True)); wb.close()
        for r in rows[1:]:
            pais = str(r[1] or '').strip().upper()
            ref_raw = str(r[3] or '').strip()
            m = re.match(r'^([A-Z]{2})(\d+)$', ref_raw)
            if m:
                country = m.group(1) if m.group(1) in VALID_COUNTRIES else pais
                ref_num = str(int(m.group(2)))
            else:
                m2 = re.search(r'(\d{3,6})$', ref_raw)
                ref_num = str(int(m2.group(1))) if m2 else ref_raw
                country = pais if pais in VALID_COUNTRIES else 'ES'
            if ref_num: sku_countries.setdefault(ref_num, set()).add(country)
    return sku_countries

def parse_sinstocks(path):
    """
    New sinstock.xlsx format — clean headers:
      Col D (3)  = EXPEDICIÓN
      Col J (9)  = ARTÍCULO  → "05993 - NOMBRE" or "A01_EU01_106744 - NOMBRE"
      Col I (8)  = CANTIDAD
      Col N (13) = ENTIDAD (marketplace/canal)
      Col S (18) = CÓDIGO DE PAÍS
      Col Z (25) = ATENCIÓN DE (not used)
    SKU extraction:
      - "05993 - ..."        → SKU = 5993  (5-digit numeric, strip leading zeros)
      - "A01_EU01_106744 -"  → SKU = A01_EU01_106744  (keep full A-prefix ref)
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Detect header row (first row with ARTÍCULO or EXPEDICIÓN)
    header_idx = 0
    for i, r in enumerate(rows[:5]):
        vals = [str(v or '').upper() for v in r]
        if any('ARTÍCULO' in v or 'ARTICULO' in v or 'EXPEDICIÓN' in v for v in vals):
            header_idx = i
            break

    headers = [str(v or '').strip().upper() for v in rows[header_idx]]

    def col(name_fragments):
        """Find column index by partial name match."""
        for frag in name_fragments:
            for i, h in enumerate(headers):
                if frag.upper() in h:
                    return i
        return None

    c_expedi    = col(['EXPEDICIÓN','EXPEDICION']) or 4
    c_articulo  = col(['ARTÍCULO','ARTICULO'])     or 9
    c_cantidad  = col(['CANTIDAD'])                or 8
    c_entidad   = col(['ENTIDAD'])                 or 13
    c_pais      = col(['CÓDIGO DE PAÍS','PAIS','PAÍS']) or 18
    c_order_id  = col(['IDENTIFICADOR'])           or 1

    records = []
    for r in rows[header_idx + 1:]:
        expedi = str(r[c_expedi] or '').strip()
        if not expedi or expedi in ('None', ''):
            continue

        articulo = str(r[c_articulo] or '').strip()
        if not articulo:
            continue

        # Parse SKU and name from ARTÍCULO
        m_num = re.match(r'^(\d{3,6})\s*[-–]\s*(.*)$', articulo)
        m_axx = re.match(r'^(A\d{2}_\w+_\d+)\s*[-–]\s*(.*)$', articulo)

        if m_num:
            sku  = str(int(m_num.group(1)))   # strip leading zeros
            name = m_num.group(2).strip()
        elif m_axx:
            sku  = m_axx.group(1).strip()     # keep full A-prefix ref
            name = m_axx.group(2).strip()
        else:
            sku  = articulo
            name = articulo

        marketplace = str(r[c_entidad] or '').strip()
        country     = str(r[c_pais]    or '').strip().upper()
        if country not in VALID_COUNTRIES:
            country = 'ES'

        is_amazon = any(k in marketplace.lower() for k in AMAZON_KEYWORDS)

        try:    qty = int(r[c_cantidad] or 1)
        except: qty = 1

        records.append({
            'EXPEDICION':  expedi,
            'SKU':         sku,
            'NOMBRE':      name[:80],
            'CANTIDAD':    qty,
            'PAIS':        country,
            'MARKETPLACE': marketplace,
            'IS_AMAZON':   is_amazon,
            'ORDER_ID':    r[c_order_id],
        })
    return pd.DataFrame(records)

# ── Tarifa lookup ──────────────────────────────────────────────────────────────
def get_tarifa(sku, country, is_amazon, tarifa_nac, tarifa_inter):
    """Returns (row, df, pvp_col, neto_col) for a given SKU and context."""
    if is_amazon and country != 'ES' and country in INTER_SHEET:
        sh, pvp_col, neto_col = INTER_SHEET[country]
        df = tarifa_inter.get(sh, tarifa_nac)
        # Fallback col names
        if pvp_col not in df.columns:
            pvp_col = next((c for c in df.columns if 'PVP PUB' in str(c).upper()), pvp_col)
        if neto_col not in df.columns:
            neto_col = next((c for c in df.columns if 'NETO' in str(c).upper()), 'NETO')
    else:
        df = tarifa_nac
        pvp_col  = 'PVP PUB.'
        neto_col = 'NETO'

    mask = df['REF'] == str(sku)
    if not mask.any():
        return None, df, pvp_col, neto_col
    return df[mask].iloc[0], df, pvp_col, neto_col

def get_stock_for_ref(sku, country, is_amazon, stocks):
    search = [country, 'ES'] if (is_amazon and country in stocks) else ['ES']
    for c in search:
        df = stocks.get(c)
        if df is None: continue
        mask = df['REF'] == str(sku)
        if mask.any():
            return int(df[mask].iloc[0]['STOCK'] or 0)
    return 0

def find_substitutes(tar_row, pvp_col, pvp_orig, sku_orig,
                     country, is_amazon, tarifa_nac, tarifa_inter, stocks, max_extra=10.0):
    """Same SUBFAMILIA, not desposicionado, PVP PUB ≤ original+max_extra, stock > 0."""
    if is_amazon and country != 'ES' and country in INTER_SHEET:
        sh, pc, nc = INTER_SHEET[country]
        df = tarifa_inter.get(sh, tarifa_nac).copy()
        if pc not in df.columns:
            pc = next((c for c in df.columns if 'PVP PUB' in str(c).upper()), pc)
        stock_country = country
    else:
        df = tarifa_nac.copy()
        pc = 'PVP PUB.'
        stock_country = 'ES'

    subfamilia = str(tar_row.get('SUBFAMILIA',''))
    familia    = str(tar_row.get('FAMILIA',''))

    # Filter by subfamilia (or familia fallback)
    if subfamilia and 'SUBFAMILIA' in df.columns:
        mask = df['SUBFAMILIA'].str.lower() == subfamilia.lower()
    elif familia and 'FAMILIA' in df.columns:
        mask = df['FAMILIA'].str.lower() == familia.lower()
    else:
        return pd.DataFrame()

    # Not desposicionado
    if 'DESPOSICIONADO' in df.columns:
        mask &= df['DESPOSICIONADO'].fillna(False) == False

    # PVP ≤ original + max_extra
    if pc in df.columns:
        pvp_vals = pd.to_numeric(df[pc], errors='coerce')
        mask &= pvp_vals <= (pvp_orig + max_extra)

    # Exclude original SKU
    mask &= df['REF'] != str(sku_orig)

    candidates = df[mask].copy()
    if candidates.empty:
        return pd.DataFrame()

    # Check stock
    stock_df  = stocks.get(stock_country, stocks.get('ES', pd.DataFrame()))
    stock_map = {} if stock_df.empty else stock_df.set_index('REF')['STOCK'].to_dict()

    results = []
    for _, row in candidates.iterrows():
        ref_c = str(row.get('REF',''))
        stk   = int(float(stock_map.get(ref_c, 0) or 0))
        if stk > 0:
            pvp_c  = float(pd.to_numeric(row.get(pc, 0), errors='coerce') or 0)
            neto_c = float(pd.to_numeric(row.get('NETO' if 'NETO' in row.index else pc, 0), errors='coerce') or 0)
            results.append({
                'REFERENCIA':      row.get('REFERENCIA',''),
                'REF':             ref_c,
                'NOMBRE COMPLETO': row.get('NOMBRE COMPLETO',''),
                'SUBFAMILIA':      row.get('SUBFAMILIA',''),
                'PVP':             pvp_c,
                'ΔPVP':            round(pvp_c - pvp_orig, 2),
                'STOCK':           stk,
            })

    if not results:
        return pd.DataFrame()
    return pd.DataFrame(results).sort_values(['ΔPVP','PVP'])

# ── Excel output ───────────────────────────────────────────────────────────────
def build_regen_excel(output_rows):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = 'Regeneracion'
    thin = Side(style='thin', color='FFD0D0CC')
    brd  = Border(top=thin, left=thin, right=thin, bottom=thin)
    for ci, h in enumerate(['NUMBER','ARTICLE','NEW_ARTICLE'], 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = Font(bold=True, color='FFFFFFFF', name='Arial', size=10)
        c.fill = PatternFill('solid', start_color='FF141413')
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = brd
    ws.row_dimensions[1].height = 22
    for ri, row in enumerate(output_rows, 2):
        fill = PatternFill('solid', start_color='FFF5F5F3' if ri%2==0 else 'FFFFFFFF')
        for ci, val in enumerate([row['NUMBER'], row['ARTICLE'], row['NEW_ARTICLE']], 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = Font(name='Arial', size=10); c.fill = fill
            c.border = brd; c.alignment = Alignment(vertical='center')
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    ws.freeze_panes = 'A2'
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── File helpers ───────────────────────────────────────────────────────────────
def _find(key, *patterns):
    if f'path_{key}' in st.session_state:
        p = st.session_state[f'path_{key}']
        if Path(p).exists(): return p
    for pat in patterns:
        m = glob.glob(str(UPLOAD_DIR / f'*{pat}*'))
        if m: return m[0]
    return None

# ── SIDEBAR uploaders ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 Cargar ficheros")
    fu_nac     = st.file_uploader('Tarifa Nacional',       type=['xlsx'], key='u_nac')
    fu_inter   = st.file_uploader('Tarifa Internacional',  type=['xlsx'], key='u_inter')
    fu_stock   = st.file_uploader('Stock Global',          type=['xlsx'], key='u_stock')
    fu_sins    = st.file_uploader('Sinstocks',             type=['xlsx'], key='u_sins')
    fu_listing = st.file_uploader('Listing Amazon',        type=['xlsx','txt','tsv','csv'], key='u_list')

    if st.button('💾 Guardar', type='primary', use_container_width=True):
        tmp = Path(tempfile.mkdtemp()); saved = 0
        for key, fu in [('nac',fu_nac),('inter',fu_inter),('stock',fu_stock),
                        ('sins',fu_sins),('listing',fu_listing)]:
            if fu:
                p = tmp / fu.name; p.write_bytes(fu.read())
                st.session_state[f'path_{key}'] = str(p); saved += 1
        for fn in [load_tarifa_nac, load_tarifa_inter, load_stock, load_listing]:
            fn.clear()
        st.success(f'✅ {saved} fichero(s) guardados')
        st.rerun()
    st.divider()
    st.caption('O coloca los ficheros en la carpeta del proyecto.')

# ── Resolve paths ──────────────────────────────────────────────────────────────
path_nac     = _find('nac',     'TARIFA_NACIONAL')
path_inter   = _find('inter',   'TARIFA_TURACO')
path_stock   = _find('stock',   'Stock_Global')
path_sins    = _find('sins',    'Sinstocks', 'sinstocks')
path_listing = _find('listing', 'Informe_', 'listing', 'Listing')
if not path_listing:
    hits = glob.glob(str(UPLOAD_DIR/'*.txt')) + glob.glob(str(UPLOAD_DIR/'*listing*'))
    if hits: path_listing = hits[0]

# Status
cs = st.columns(5)
for col, lbl, pth in zip(cs,
    ['Tarifa Nac.','Tarifa Inter','Stock Global','Sinstocks','Listing Amazon'],
    [path_nac, path_inter, path_stock, path_sins, path_listing]):
    with col:
        (st.success if pth else st.warning)(f'{"✅" if pth else "⚠️"} {lbl}')

data_ok = all([path_nac, path_inter, path_stock, path_sins])

# ── TABS ───────────────────────────────────────────────────────────────────────
tab_proc, tab_res = st.tabs(['⚙️ Procesar', '📊 Resultados'])

with tab_proc:
    if not data_ok:
        st.warning('👈 Sube los ficheros en el panel izquierdo para continuar.')
    else:
        tarifa_nac   = load_tarifa_nac(path_nac)
        tarifa_inter = load_tarifa_inter(path_inter)
        stocks       = load_stock(path_stock)
        listing_map  = load_listing(path_listing) if path_listing else {}

        st.markdown('<div class="sec">⚙️ Configuración</div>', unsafe_allow_html=True)
        ca, cb, cc = st.columns(3)
        with ca: max_extra       = st.number_input('Máx. incremento PVP PUB (€)', 0.0, 50.0, 10.0, 1.0)
        with cb: default_country = st.selectbox('País por defecto', ['ES','FR','IT','DE','PT','PL'], index=0)
        with cc: skip_dupes      = st.checkbox('Ignorar DUPLICADO', value=True)

        if st.button('🚀 Procesar pedidos sin stock', type='primary', use_container_width=True):
            with st.spinner('Leyendo Sinstocks…'):
                df_sins = parse_sinstocks(path_sins)
                if skip_dupes:
                    df_sins = df_sins[~df_sins['EXPEDICION'].str.upper().str.startswith('DUPLICADO')]
                df_sins = df_sins[df_sins['EXPEDICION'].str.startswith('D')].copy()
                # Enrich country from listing for Amazon orders
                if listing_map:
                    def enrich(row):
                        if row['IS_AMAZON'] and row['PAIS'] == 'ES':
                            ctrs = listing_map.get(str(row['SKU']), set()) - {'ES'}
                            if len(ctrs) == 1: return list(ctrs)[0]
                        return row['PAIS']
                    df_sins['PAIS'] = df_sins.apply(enrich, axis=1)
                df_sins['PAIS'] = df_sins['PAIS'].fillna(default_country)

            prog = st.progress(0)
            total = len(df_sins)
            results = []

            for i, (_, row) in enumerate(df_sins.iterrows()):
                prog.progress((i+1)/total,
                    text=f"[{i+1}/{total}] {row['EXPEDICION']} · {row['NOMBRE'][:40]}")

                sku     = row['SKU']
                country = str(row['PAIS'])
                is_amz  = bool(row['IS_AMAZON'])

                tar_row, df_tar, pvp_col, neto_col = get_tarifa(
                    sku, country, is_amz, tarifa_nac, tarifa_inter)

                pvp_orig   = 0.0
                subfamilia = ''
                nombre_orig= row['NOMBRE']

                if tar_row is not None:
                    try: pvp_orig = float(pd.to_numeric(tar_row.get(pvp_col,0), errors='coerce') or 0)
                    except: pass
                    subfamilia  = str(tar_row.get('SUBFAMILIA',''))
                    nombre_orig = str(tar_row.get('NOMBRE COMPLETO', row['NOMBRE']))

                subs = pd.DataFrame()
                if tar_row is not None and pvp_orig > 0:
                    subs = find_substitutes(tar_row, pvp_col, pvp_orig, sku,
                                            country, is_amz, tarifa_nac, tarifa_inter,
                                            stocks, max_extra)

                results.append({
                    'row':         row.to_dict(),
                    'tar_row':     tar_row,
                    'pvp_col':     pvp_col,
                    'pvp_orig':    pvp_orig,
                    'subfamilia':  subfamilia,
                    'nombre_orig': nombre_orig,
                    'sku':         sku,
                    'subs':        subs,
                })

            prog.empty()
            st.session_state['results'] = results
            st.toast('✅ Procesado completado', icon='🎯')
            st.components.v1.html("""
            <script>
            function click(){const t=window.parent.document.querySelectorAll('[data-baseweb="tab"]');
            for(const x of t){if(x.textContent.includes('Resultados')){x.click();return true;}}return false;}
            let n=0,iv=setInterval(()=>{if(click()||++n>20)clearInterval(iv);},150);
            </script>""", height=0)
            st.rerun()

with tab_res:
    if 'results' not in st.session_state:
        st.info('Ejecuta el procesado en ⚙️ Procesar.')
        st.stop()

    results  = st.session_state['results']
    con_sust = [r for r in results if not r['subs'].empty]
    sin_sust = [r for r in results if r['subs'].empty]
    total    = len(results)

    st.markdown(f"""<div class="kpi-row">
      <div class="kpi"><div class="val">{total}</div><div class="lbl">Pedidos sin stock</div></div>
      <div class="kpi"><div class="val">{len(con_sust)}</div><div class="lbl">Con sustituto ✅</div></div>
      <div class="kpi"><div class="val">{len(sin_sust)}</div><div class="lbl">A cancelar ❌</div></div>
      <div class="kpi"><div class="val">{len(con_sust)*100//total if total else 0}%</div><div class="lbl">Tasa resolución</div></div>
    </div>""", unsafe_allow_html=True)

    t1, t2 = st.tabs(['✅ Con sustituto', '❌ A cancelar'])

    with t1:
        if not con_sust:
            st.info('No hay pedidos con sustituto disponible.')
        else:
            table_rows, output_rows = [], []
            for r in con_sust:
                row  = r['row']
                best = r['subs'].iloc[0]
                delta = float(best.get('ΔPVP', 0) or 0)
                table_rows.append({
                    'Expedición':        row['EXPEDICION'],
                    'País':              row['PAIS'],
                    'Canal':             row['MARKETPLACE'][:30],
                    'Amazon':            '✅' if row['IS_AMAZON'] else '—',
                    'SKU original':      row['SKU'],
                    'Producto original': r['nombre_orig'][:55],
                    'PVP orig. (€)':     round(r['pvp_orig'], 2),
                    'Subfamilia':        r['subfamilia'],
                    '→ SKU sustituto':   str(best.get('REFERENCIA','')),
                    '→ Sustituto':       str(best.get('NOMBRE COMPLETO',''))[:55],
                    'PVP sust. (€)':     round(float(best.get('PVP',0) or 0), 2),
                    'Δ PVP (€)':         delta,
                    'Stock disponible':  int(best.get('STOCK',0) or 0),
                })
                output_rows.append({
                    'NUMBER':      row['EXPEDICION'],
                    'ARTICLE':     row['SKU'],
                    'NEW_ARTICLE': str(best.get('REFERENCIA','')),
                })

            df_t = pd.DataFrame(table_rows)
            st.dataframe(df_t, use_container_width=True, hide_index=True, column_config={
                'PVP orig. (€)': st.column_config.NumberColumn(format='%.2f €'),
                'PVP sust. (€)': st.column_config.NumberColumn(format='%.2f €'),
                'Δ PVP (€)':     st.column_config.NumberColumn(format='%.2f €'),
            })

            dc1, dc2 = st.columns(2)
            with dc1:
                st.download_button('⬇️ Plantilla regeneración Excel',
                    build_regen_excel(output_rows), 'regeneracion_pedidos.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True, type='primary')
            with dc2:
                st.download_button('⬇️ CSV sustitutos',
                    df_t.to_csv(index=False).encode(), 'sustitutos.csv',
                    'text/csv', use_container_width=True)

            st.markdown('<div class="sec">🔎 Detalle por pedido</div>', unsafe_allow_html=True)
            for r in con_sust:
                row  = r['row']
                best = r['subs'].iloc[0]
                delta = float(best.get('ΔPVP',0) or 0)
                sign  = f'+{delta:.2f}€' if delta>0 else (f'{delta:.2f}€' if delta<0 else '±0')
                with st.expander(
                    f"**{row['EXPEDICION']}** · {row['PAIS']} "
                    f"{'🛒' if row['IS_AMAZON'] else ''} · "
                    f"SKU {row['SKU']} → **{best.get('REFERENCIA','')}** ({sign} PVP)"
                ):
                    cc1, cc2 = st.columns(2)
                    with cc1:
                        st.markdown('**📦 Pedido original**')
                        st.markdown(f"Expedición: `{row['EXPEDICION']}`")
                        st.markdown(f"Canal: `{row['MARKETPLACE']}`")
                        st.markdown(f"SKU: `{row['SKU']}` · País: **{row['PAIS']}**")
                        st.markdown(f"**{r['nombre_orig']}**")
                        st.markdown(f"PVP PUB: **{r['pvp_orig']:.2f}€** · Subfamilia: {r['subfamilia']}")
                    with cc2:
                        st.markdown('**🔄 Sustitutos con stock**')
                        show = r['subs'][['REFERENCIA','NOMBRE COMPLETO','PVP','ΔPVP','STOCK']].rename(
                            columns={'NOMBRE COMPLETO':'Nombre','PVP':'PVP PUB (€)','ΔPVP':'Δ PVP'})
                        st.dataframe(show.head(5), use_container_width=True, hide_index=True,
                            column_config={
                                'PVP PUB (€)': st.column_config.NumberColumn(format='%.2f €'),
                                'Δ PVP':       st.column_config.NumberColumn(format='%.2f €'),
                            })

    with t2:
        if not sin_sust:
            st.success('✅ Todos los pedidos tienen sustituto disponible.')
        else:
            cancel_rows = []
            for r in sin_sust:
                row = r['row']
                motivo = ('SKU no encontrado en tarifa' if r['tar_row'] is None
                          else f"Sin sustituto con stock en subfamilia '{r['subfamilia']}'")
                cancel_rows.append({
                    'Expedición':   row['EXPEDICION'],
                    'País':         row['PAIS'],
                    'Canal':        row['MARKETPLACE'][:35],
                    'Amazon':       '✅' if row['IS_AMAZON'] else '—',
                    'SKU':          row['SKU'],
                    'Producto':     r['nombre_orig'][:60],
                    'PVP (€)':      round(r['pvp_orig'],2),
                    'Subfamilia':   r['subfamilia'],
                    'Motivo':       motivo,
                })
            df_c = pd.DataFrame(cancel_rows)
            st.warning(f"⚠️ **{len(df_c)} pedidos** sin sustituto — deben cancelarse.")
            st.dataframe(df_c, use_container_width=True, hide_index=True,
                column_config={'PVP (€)': st.column_config.NumberColumn(format='%.2f €')})
            st.download_button('⬇️ CSV cancelaciones',
                df_c.to_csv(index=False).encode(), 'cancelaciones.csv',
                'text/csv', use_container_width=True)
