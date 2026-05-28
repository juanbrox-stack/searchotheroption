"""
Gestor de Sustitutos Cecotec
Pedidos sin stock → busca sustitutos con stock en misma subfamilia
Regla: si Amazon → tarifa inter del país; resto → tarifa nacional + stock España
"""
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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
VALID_COUNTRIES = {'ES','FR','IT','DE','PT','PL','NL','BE','UK'}
AMAZON_KEYWORDS = ['amazon', 'turaco-amazon', 'mano-a-mano']

# Tarifa inter sheet per country
INTER_SHEET = {
    'FR': ('ES-FR', 'NETO ES-FR'),
    'IT': ('ES-IT', 'NETO ES-IT'),
    'DE': ('ES-DE', 'NETO ES-DE'),
    'PT': ('PT',    'NETO PT'),
    'PL': ('PL',    'NETO PL'),
    'NL': ('NL',    'NETO NL'),
    'BE': ('BE',    'NETO BE'),
}

UPLOAD_DIR = Path(__file__).parent

# ── Data loaders ───────────────────────────────────────────────────────────────
@st.cache_data
def load_tarifa_nac(path):
    df = pd.read_excel(path, sheet_name='T_AMZ', dtype={'REFERENCIA': str})
    df['REF'] = df['REFERENCIA'].apply(lambda x: str(int(float(x))) if re.match(r'^\d+\.?\d*$', str(x)) else str(x))
    return df

@st.cache_data
def load_tarifa_inter(path):
    xl = pd.ExcelFile(path)
    out = {}
    for sh in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sh, dtype={'REFERENCIA': str})
        # Some sheets may not have REFERENCIA (e.g. % sheets) — skip them
        if 'REFERENCIA' not in df.columns:
            continue
        df['REF'] = df['REFERENCIA'].apply(
            lambda x: str(int(float(x))) if re.match(r'^\d+\.?\d*$', str(x).strip()) else str(x)
        )
        out[sh] = df
    return out

@st.cache_data
def load_stock_global(path):
    xl = pd.ExcelFile(path)
    SHEET_MAP = {'España':'ES','Alemania':'DE','Francia':'FR','Italia':'IT'}
    out = {}
    for sh in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sh, dtype={'Referencia': str})
        df['REF'] = df['Referencia'].apply(lambda x: re.search(r'(\d{3,6})$', str(x)).group(1).lstrip('0') or
                                           re.search(r'(\d{3,6})$', str(x)).group(1)
                                           if re.search(r'(\d{3,6})$', str(x)) else str(x))
        df['REF'] = df['REF'].apply(lambda x: str(int(x)) if re.match(r'^\d+$', str(x)) else str(x))
        out[SHEET_MAP.get(sh, sh)] = df
    return out

@st.cache_data
def load_listing(path):
    """Amazon listing/orders: SKU → countries.
    Supports:
    - Orders TSV (amazon-order-id + sku + ship-country columns)
    - Listings TSV (SKU del vendedor column with FR/IT/DE prefixes)
    - Legacy xlsx (Informe_W20 format)
    """
    sku_countries = {}
    p = Path(path)

    if p.suffix.lower() in ('.txt', '.tsv', '.csv'):
        df = pd.read_csv(path, sep='\t', encoding='utf-8-sig', dtype=str)
        cols = [c.strip() for c in df.columns]
        df.columns = cols

        # ── Amazon orders format: has 'sku' and 'ship-country' ──────────────
        if 'sku' in cols and 'ship-country' in cols:
            for _, row in df[['sku','ship-country']].dropna().iterrows():
                sku = str(row['sku']).strip()
                country = str(row['ship-country']).strip().upper()
                if country not in VALID_COUNTRIES:
                    continue
                # S08303 → 8303 | S00120 → 120 | A01_EU01_008303 → 8303
                m = re.search(r'(\d{3,6})$', sku)
                if m:
                    ref_num = str(int(m.group(1)))
                    sku_countries.setdefault(ref_num, set()).add(country)

        # ── Listings format: has 'SKU del vendedor' with FR/IT/DE prefix ────
        elif 'SKU del vendedor' in cols:
            for sku_raw in df['SKU del vendedor'].dropna():
                sku = str(sku_raw).strip()
                m = re.match(r'^([A-Z]{2})0*(\d+)$', sku)
                if m and m.group(1) in VALID_COUNTRIES:
                    sku_countries.setdefault(str(int(m.group(2))), set()).add(m.group(1))
                else:
                    m2 = re.search(r'(\d{3,6})$', sku)
                    if m2:
                        sku_countries.setdefault(str(int(m2.group(1))), set()).add('ES')

    else:
        # ── Legacy xlsx (Informe_W20) ────────────────────────────────────────
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb['BBDD'] if 'BBDD' in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
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
            if ref_num:
                sku_countries.setdefault(ref_num, set()).add(country)

    return sku_countries

def parse_sinstocks(path):
    """Parse Sinstocks with robust column detection."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    records = []
    for r in rows[1:]:
        expedi = str(r[4] or '').strip()
        if not expedi or expedi == 'None':
            continue

        # ── SKU: first integer in cols 12-22 that looks like a ref ──
        ref = None
        for i in range(12, 23):
            v = r[i]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                n = int(v)
                if 100 <= n <= 99999 and n not in range(2020, 2031):
                    ref = str(n)
                    break
            elif isinstance(v, str) and re.match(r'^[A-Z]\d{2}_[A-Z]{2}\d{2}_\d{5,}$', v.strip()):
                ref = v.strip()
                break

        # ── Product name: join non-null string cells cols 12-20 ──
        name_parts = []
        for i in range(12, 21):
            v = r[i]
            if isinstance(v, str) and v.strip() not in ('None',''):
                name_parts.append(v.strip())
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                name_parts.append(str(int(v)))
        product_name = ' '.join(name_parts[:5])

        # ── Marketplace: scan all values for known keywords ──
        marketplace = ''
        for i in range(6, len(r)):
            v = str(r[i] or '').strip()
            if any(k in v.lower() for k in ['turaco','amazon','cdiscount','miravia','temu',
                                              'shein','mediamarkt','darty','showroom','mirakl',
                                              'worten','carrefour','mano-a-mano','arise','tiktok',
                                              'rakuten','fnac','leroy','veepee']):
                marketplace = v
                break

        # ── Country: scan right-to-left for 2-letter country code ──
        country = 'ES'
        for i in range(len(r)-1, 10, -1):
            v = str(r[i] or '').strip().upper()
            if v in VALID_COUNTRIES:
                country = v
                break

        # ── Is Amazon order? ──
        is_amazon = any(k in marketplace.lower() for k in AMAZON_KEYWORDS)

        # ── Price ──
        precio = None
        for i in range(36, 39):
            try:
                v = r[i]
                if v and str(v) not in ('None','CECOTEC',''):
                    precio = float(str(v).replace(',','.'))
                    break
            except Exception:
                pass

        records.append({
            'EXPEDICION': expedi,
            'REF': ref,
            'NOMBRE': product_name[:80],
            'PAIS': country,
            'MARKETPLACE': marketplace,
            'IS_AMAZON': is_amazon,
            'PRECIO_VENTA': precio,
            'ORDER_ID': r[1],
        })
    return pd.DataFrame(records)

# ── Business logic ─────────────────────────────────────────────────────────────
def get_tarifa_row(ref, country, is_amazon, tarifa_nac, tarifa_inter):
    """Get the tarifa row and neto_col for a given ref and context."""
    if is_amazon and country != 'ES' and country in INTER_SHEET:
        sheet, neto_col = INTER_SHEET[country]
        df = tarifa_inter.get(sheet, tarifa_nac)
        if neto_col not in df.columns:
            neto_col = next((c for c in df.columns if 'NETO' in c.upper()), 'NETO')
    else:
        df = tarifa_nac
        neto_col = 'NETO'

    mask = df['REF'] == str(ref)
    if not mask.any():
        return None, df, neto_col
    return df[mask].iloc[0], df, neto_col

def get_stock(ref, country, is_amazon, stocks):
    """Get stock operativo for a ref. Amazon+foreign → country stock; else → ES."""
    search_countries = [country, 'ES'] if (is_amazon and country in stocks) else ['ES']
    for c in search_countries:
        df = stocks.get(c)
        if df is None: continue
        m = df['REF'] == str(ref)
        if m.any():
            return int(df[m].iloc[0].get('Stock Operativo', 0) or 0)
    return 0

def find_substitutes(ref_row, neto_col, neto_orig, ref_orig,
                     country, is_amazon, tarifa_nac, tarifa_inter, stocks, max_extra=10.0):
    """Find substitutes: same subfamilia, stock>0, neto ≤ orig+max_extra."""
    if is_amazon and country != 'ES' and country in INTER_SHEET:
        sheet, nc = INTER_SHEET[country]
        df = tarifa_inter.get(sheet, tarifa_nac).copy()
        if nc not in df.columns:
            nc = next((c for c in df.columns if 'NETO' in c.upper()), 'NETO')
        stock_country = country
    else:
        df = tarifa_nac.copy()
        nc = 'NETO'
        stock_country = 'ES'

    subfamilia = str(ref_row.get('SUBFAMILIA','')) if ref_row is not None else ''
    familia    = str(ref_row.get('FAMILIA',''))    if ref_row is not None else ''

    # Filter by subfamilia (or familia as fallback)
    if subfamilia and 'SUBFAMILIA' in df.columns:
        mask = df['SUBFAMILIA'].str.lower() == subfamilia.lower()
    elif familia and 'FAMILIA' in df.columns:
        mask = df['FAMILIA'].str.lower() == familia.lower()
    else:
        return pd.DataFrame()

    # Not desposicionado
    if 'DESPOSICIONADO' in df.columns:
        mask &= df['DESPOSICIONADO'].fillna(False) == False

    # Price cap: neto ≤ original + max_extra
    if nc in df.columns:
        neto_series = pd.to_numeric(df[nc], errors='coerce')
        mask &= neto_series <= (neto_orig + max_extra)

    # Exclude original ref
    mask &= df['REF'] != str(ref_orig)

    candidates = df[mask].copy()
    if candidates.empty:
        return pd.DataFrame()

    # Check stock
    stock_df = stocks.get(stock_country, stocks.get('ES', pd.DataFrame()))
    stock_map = {} if stock_df.empty else stock_df.set_index('REF')['Stock Operativo'].to_dict()

    results = []
    for _, row in candidates.iterrows():
        ref_c = str(row.get('REF',''))
        stk = int(float(stock_map.get(ref_c, 0) or 0))
        if stk > 0:
            neto_c = float(pd.to_numeric(row.get(nc, 0), errors='coerce') or 0)
            results.append({
                'REFERENCIA':     row.get('REFERENCIA',''),
                'NOMBRE COMPLETO':row.get('NOMBRE COMPLETO',''),
                'SUBFAMILIA':     row.get('SUBFAMILIA',''),
                'NETO':           neto_c,
                'ΔNETO':          round(neto_c - neto_orig, 2),
                'STOCK':          stk,
                'REF':            ref_c,
            })

    if not results:
        return pd.DataFrame()
    return pd.DataFrame(results).sort_values(['ΔNETO','NETO'])

# ── Excel output ───────────────────────────────────────────────────────────────
def build_regen_excel(output_rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Regeneracion'
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
            c.font = Font(name='Arial', size=10)
            c.fill = fill
            c.border = brd
            c.alignment = Alignment(vertical='center')

    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14
    ws.freeze_panes = 'A2'

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── File path helpers ──────────────────────────────────────────────────────────
def _find(key, *patterns):
    if f'path_{key}' in st.session_state:
        return st.session_state[f'path_{key}']
    for p in patterns:
        matches = glob.glob(str(UPLOAD_DIR / f'*{p}*'))
        if matches: return matches[0]
    return None

# ── UI ─────────────────────────────────────────────────────────────────────────
# ── File uploaders FIRST (always visible, before tabs) ────────────────────────
with st.sidebar:
    st.markdown("### 📂 Cargar ficheros")
    fu_nac     = st.file_uploader('Tarifa Nacional',     type=['xlsx'], key='u_nac')
    fu_inter   = st.file_uploader('Tarifa Internacional',type=['xlsx'], key='u_inter')
    fu_stock   = st.file_uploader('Stock Global',        type=['xlsx'], key='u_stock')
    fu_sins    = st.file_uploader('Sinstocks',           type=['xlsx'], key='u_sins')
    fu_listing = st.file_uploader('Listing Amazon',      type=['xlsx','txt','tsv','csv'], key='u_list')

    if st.button('💾 Guardar', type='primary', use_container_width=True):
        tmp = Path(tempfile.mkdtemp())
        saved = 0
        for key, fu in [('nac',fu_nac),('inter',fu_inter),('stock',fu_stock),
                        ('sins',fu_sins),('listing',fu_listing)]:
            if fu:
                p = tmp / fu.name
                p.write_bytes(fu.read())
                st.session_state[f'path_{key}'] = str(p)
                saved += 1
        for fn in [load_tarifa_nac, load_tarifa_inter, load_stock_global, load_listing]:
            fn.clear()
        st.success(f'✅ {saved} fichero(s) guardados')
        st.rerun()

    st.divider()
    st.caption('O coloca los ficheros en la misma carpeta que app.py y reinicia.')

# Resolve paths
path_nac     = _find('nac',     'TARIFA_NACIONAL')
path_inter   = _find('inter',   'TARIFA_TURACO')
path_stock   = _find('stock',   'Stock_Global')
path_sins    = _find('sins',    'Sinstocks', 'sinstocks')
path_listing = _find('listing', 'Informe_', 'listing', 'Listing')
if not path_listing:
    # Also search for .txt files
    txt_matches = glob.glob(str(UPLOAD_DIR / '*Informe*listing*.txt')) + \
                  glob.glob(str(UPLOAD_DIR / '*listing*.txt')) + \
                  glob.glob(str(UPLOAD_DIR / '*Listing*.txt'))
    if txt_matches: path_listing = txt_matches[0]

# Status bar
cs = st.columns(5)
labels = ['Tarifa Nac.','Tarifa Inter','Stock Global','Sinstocks','Listing Amazon']
paths  = [path_nac, path_inter, path_stock, path_sins, path_listing]
for col, label, path in zip(cs, labels, paths):
    with col:
        if path: st.success(f'✅ {label}')
        else:    st.warning(f'⚠️ {label}')

data_ok = all([path_nac, path_inter, path_stock, path_sins])

tab_proc, tab_res = st.tabs(['⚙️ Procesar', '📊 Resultados'])

with tab_proc:
    if not data_ok:
        st.warning('👈 Sube los ficheros en el **panel izquierdo** para continuar.')
    else:

        # Load data
        tarifa_nac   = load_tarifa_nac(path_nac)
        tarifa_inter = load_tarifa_inter(path_inter)
        stocks       = load_stock_global(path_stock)
        listing_map  = load_listing(path_listing) if path_listing else {}

    st.markdown('<div class="sec">⚙️ Configuración</div>', unsafe_allow_html=True)
    ca, cb, cc = st.columns(3)
    with ca: max_extra = st.number_input('Máx. incremento neto (€)', 0.0, 50.0, 10.0, 1.0)
    with cb: default_country = st.selectbox('País por defecto si no se detecta',
                                             ['ES','FR','IT','DE','PT','PL'], index=0)
    with cc: skip_dupes = st.checkbox('Ignorar expediciones DUPLICADO', value=True)

    if st.button('🚀 Procesar pedidos sin stock', type='primary', use_container_width=True):
        with st.spinner('Leyendo Sinstocks…'):
            df_sins = parse_sinstocks(path_sins)
            if skip_dupes:
                df_sins = df_sins[~df_sins['EXPEDICION'].str.upper().str.startswith('DUPLICADO')]
            df_sins = df_sins[df_sins['EXPEDICION'].str.startswith('D')].copy()
            # Apply listing country for Amazon orders without clear country detection
            if listing_map:
                def enrich_country(row):
                    if row['IS_AMAZON'] and row['REF'] and row['PAIS'] == 'ES':
                        countries = listing_map.get(str(row['REF']), set())
                        countries_non_es = countries - {'ES'}
                        if len(countries_non_es) == 1:
                            return list(countries_non_es)[0]
                    return row['PAIS']
                df_sins['PAIS'] = df_sins.apply(enrich_country, axis=1)
            df_sins['PAIS'] = df_sins['PAIS'].fillna(default_country)

        prog = st.progress(0)
        total = len(df_sins)
        results = []

        for i, (_, row) in enumerate(df_sins.iterrows()):
            prog.progress((i+1)/total, text=f"[{i+1}/{total}] {row['EXPEDICION']} · {row['NOMBRE'][:40]}")

            ref   = row['REF']
            country = str(row['PAIS'])
            is_amz  = bool(row['IS_AMAZON'])

            # Get tarifa row
            tar_row, df_tar, neto_col = get_tarifa_row(ref, country, is_amz, tarifa_nac, tarifa_inter)
            neto_orig = 0.0
            subfamilia = nombre_orig = ''
            if tar_row is not None:
                try: neto_orig = float(pd.to_numeric(tar_row.get(neto_col, 0), errors='coerce') or 0)
                except: pass
                subfamilia = str(tar_row.get('SUBFAMILIA',''))
                nombre_orig = str(tar_row.get('NOMBRE COMPLETO', row['NOMBRE']))

            # Stock of original
            stk_orig = get_stock(ref or '', country, is_amz, stocks) if ref else 0

            # Find substitutes
            subs = pd.DataFrame()
            if tar_row is not None and neto_orig > 0:
                subs = find_substitutes(tar_row, neto_col, neto_orig, ref or '',
                                        country, is_amz, tarifa_nac, tarifa_inter,
                                        stocks, max_extra)

            results.append({
                'row':        row.to_dict(),
                'tar_row':    tar_row,
                'neto_col':   neto_col,
                'neto_orig':  neto_orig,
                'subfamilia': subfamilia,
                'nombre_orig':nombre_orig or row['NOMBRE'],
                'stk_orig':   stk_orig,
                'subs':       subs,
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

    st.markdown(f"""
    <div class="kpi-row">
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
                delta = float(best.get('ΔNETO', 0) or 0)
                table_rows.append({
                    'Expedición':         row['EXPEDICION'],
                    'País':               row['PAIS'],
                    'Canal':              row['MARKETPLACE'][:30],
                    'Amazon':             '✅' if row['IS_AMAZON'] else '—',
                    'Ref. original':      row['REF'],
                    'Producto original':  r['nombre_orig'][:55],
                    'Neto orig. (€)':     round(r['neto_orig'],2),
                    'Subfamilia':         r['subfamilia'],
                    '→ Ref. sustituto':   str(best.get('REFERENCIA','')),
                    '→ Sustituto':        str(best.get('NOMBRE COMPLETO',''))[:55],
                    'Neto sust. (€)':     round(float(best.get('NETO',0) or 0),2),
                    'Δ Neto (€)':         delta,
                    'Stock disponible':   int(best.get('STOCK',0) or 0),
                })
                output_rows.append({
                    'NUMBER':      row['EXPEDICION'],
                    'ARTICLE':     row['REF'] or '',
                    'NEW_ARTICLE': str(best.get('REFERENCIA','')),
                })

            df_t = pd.DataFrame(table_rows)
            st.dataframe(df_t, use_container_width=True, hide_index=True, column_config={
                'Neto orig. (€)': st.column_config.NumberColumn(format='%.2f €'),
                'Neto sust. (€)': st.column_config.NumberColumn(format='%.2f €'),
                'Δ Neto (€)':     st.column_config.NumberColumn(format='%.2f €'),
            })

            dc1, dc2 = st.columns(2)
            with dc1:
                st.download_button('⬇️ Plantilla regeneración Excel', build_regen_excel(output_rows),
                    'regeneracion_pedidos.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True, type='primary')
            with dc2:
                st.download_button('⬇️ CSV sustitutos', df_t.to_csv(index=False).encode(),
                    'sustitutos.csv', 'text/csv', use_container_width=True)

            st.markdown('<div class="sec">🔎 Detalle por pedido</div>', unsafe_allow_html=True)
            for r in con_sust:
                row  = r['row']
                best = r['subs'].iloc[0]
                delta = float(best.get('ΔNETO',0) or 0)
                sign  = f'+{delta:.2f}€' if delta > 0 else (f'{delta:.2f}€' if delta < 0 else '±0')
                with st.expander(
                    f"**{row['EXPEDICION']}** · {row['PAIS']} {'🛒' if row['IS_AMAZON'] else ''} · "
                    f"Ref {row['REF']} → **{best.get('REFERENCIA','')}** ({sign} neto)"
                ):
                    cc1, cc2 = st.columns(2)
                    with cc1:
                        st.markdown('**📦 Pedido original**')
                        st.markdown(f"Expedición: `{row['EXPEDICION']}`")
                        st.markdown(f"Canal: `{row['MARKETPLACE']}`")
                        st.markdown(f"Ref: `{row['REF']}` · País: **{row['PAIS']}**"
                                    + (' · Amazon' if row['IS_AMAZON'] else ''))
                        st.markdown(f"**{r['nombre_orig']}**")
                        st.markdown(f"Neto: **{r['neto_orig']:.2f}€** · Subfamilia: {r['subfamilia']}")
                    with cc2:
                        st.markdown('**🔄 Sustitutos con stock**')
                        show = r['subs'][['REFERENCIA','NOMBRE COMPLETO','NETO','ΔNETO','STOCK']].rename(
                            columns={'NOMBRE COMPLETO':'Nombre','NETO':'Neto (€)','ΔNETO':'Δ Neto'})
                        st.dataframe(show.head(5), use_container_width=True, hide_index=True,
                            column_config={
                                'Neto (€)': st.column_config.NumberColumn(format='%.2f €'),
                                'Δ Neto':   st.column_config.NumberColumn(format='%.2f €'),
                            })

    with t2:
        if not sin_sust:
            st.success('✅ Todos los pedidos tienen sustituto disponible.')
        else:
            cancel_rows = []
            for r in sin_sust:
                row = r['row']
                motivo = ('Sin referencia detectada' if not row['REF']
                          else 'Sin referencia en tarifa' if r['tar_row'] is None
                          else f"Sin sustituto con stock en subfamilia '{r['subfamilia']}'")
                cancel_rows.append({
                    'Expedición':   row['EXPEDICION'],
                    'País':         row['PAIS'],
                    'Canal':        row['MARKETPLACE'][:35],
                    'Amazon':       '✅' if row['IS_AMAZON'] else '—',
                    'Ref.':         row['REF'],
                    'Producto':     r['nombre_orig'][:60],
                    'Neto (€)':     round(r['neto_orig'],2),
                    'Subfamilia':   r['subfamilia'],
                    'Motivo':       motivo,
                })
            df_c = pd.DataFrame(cancel_rows)
            st.warning(f"⚠️ **{len(df_c)} pedidos** sin sustituto disponible.")
            st.dataframe(df_c, use_container_width=True, hide_index=True,
                column_config={'Neto (€)': st.column_config.NumberColumn(format='%.2f €')})
            st.download_button('⬇️ CSV cancelaciones', df_c.to_csv(index=False).encode(),
                'cancelaciones.csv', 'text/csv', use_container_width=True)
