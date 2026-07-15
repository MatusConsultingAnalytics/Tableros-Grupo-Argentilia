import pandas as pd
import os
import json
import shutil
import gdown
from datetime import datetime, date

# ── Timestamp ────────────────────────────────────────────────────────
ULTIMA_ACTUALIZACION = datetime.now().strftime("%d/%m/%Y %I:%M %p")

# ── Google Drive ─────────────────────────────────────────────────────
FOLDER_ID = "1RtARlWhIZNG2cyxFy8Bna6R5uOdoCS97"
CARPETA   = "/tmp/argentilia_excel"

ARCHIVOS = {
    "Argentilia León":      "ARGENTILIA  LEON .xlsx",
    "Argentilia Querétaro": "ARGENTILIA  QRO.xlsx",
    "Frascati":             "FRASCATI.xlsx",
    "Mikoh":                "MIKOH.xlsx",
}

# Unidades con columna única "Comensales" (vs H/M/N)
COMENSALES_SIMPLE = {"Argentilia León", "Mikoh"}

MESES_ORDEN = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
MESES_ES    = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

DIAS_SEMANA = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]

# ── Descarga ──────────────────────────────────────────────────────────
def descargar_archivos():
    if os.path.exists(CARPETA):
        shutil.rmtree(CARPETA)
    print("⬇️  Descargando archivos desde Google Drive...")
    gdown.download_folder(id=FOLDER_ID, output=CARPETA, quiet=False, use_cookies=False)

# ── Helpers ───────────────────────────────────────────────────────────
def safe_float(v):
    try:
        if pd.isna(v): return 0.0
        s = str(v).strip().replace('$','').replace(',','').replace(' ','')
        if s in ['-','$ -','nan','']: return 0.0
        return float(s)
    except: return 0.0

def parse_fecha(v):
    if pd.isna(v): return None
    s = str(v).strip()
    for fmt in ('%d-%b-%Y','%Y-%m-%d','%d/%m/%Y','%m/%d/%Y'):
        try: return datetime.strptime(s, fmt).date()
        except: pass
    try: return pd.to_datetime(v).date()
    except: return None

# ── Lectura de hoja principal ─────────────────────────────────────────
def leer_excel(path, nombre):
    df = pd.read_excel(path, sheet_name=0, header=None)
    simple = nombre in COMENSALES_SIMPLE

    meses_data  = {}   # mes -> { semanas, total, cheque, clientes, mix_ali, mix_beb, dias }
    current_mes = None
    presup_row  = None
    cap_start   = None
    cap_end     = None
    in_cap      = False
    in_resumen  = False

    i = 0
    while i < len(df):
        row  = df.iloc[i]
        c0   = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        c0up = c0.upper().replace(' ','')

        # ── Detectar inicio de mes ────────────────────────────────────
        for idx, mk in enumerate(MESES_ORDEN):
            if c0up == mk:
                if current_mes and cap_start is not None:
                    cap_end = i - 1
                current_mes = MESES_ES[idx]
                meses_data[current_mes] = {
                    'semanas': [], 'total': None, 'cheque': None,
                    'clientes': None, 'mix_ali': None, 'mix_beb': None, 'dias': []
                }
                presup_row = None
                cap_start  = None
                cap_end    = None
                in_cap     = False
                in_resumen = False
                break

        if not current_mes:
            i += 1
            continue

        d = meses_data[current_mes]

        # ── Presupuesto semanal ───────────────────────────────────────
        if c0 == 'Meta de venta ($)':
            presup_row = i
            i += 1
            continue

        # ── Inicio de captura diaria (detectar por encabezado Fecha) ─
        if c0 == 'Fecha' and not in_resumen:
            in_cap    = True
            cap_start = i + 1
            i += 1
            continue

        # ── Datos diarios ─────────────────────────────────────────────
        if in_cap and not in_resumen:
            fecha = parse_fecha(row.iloc[0])
            if fecha:
                ali = safe_float(row.iloc[1])
                beb = safe_float(row.iloc[2])
                if simple:
                    com = safe_float(row.iloc[3])
                else:
                    h = safe_float(row.iloc[3])
                    m = safe_float(row.iloc[4])
                    n = safe_float(row.iloc[5])
                    com = h + m + n
                d['dias'].append({
                    'fecha': fecha,
                    'dow': fecha.weekday(),          # 0=Lunes … 6=Domingo
                    'alimentos': ali,
                    'bebidas': beb,
                    'total': ali + beb,
                    'comensales': com
                })
            # Detectar fin de bloque diario
            if c0 == 'RESUMEN AUTOMÁTICO (no editar — se calcula solo)':
                in_cap     = False
                in_resumen = True
                cap_end    = i - 1
            i += 1
            continue

        # ── Resumen automático ────────────────────────────────────────
        if in_resumen:
            # Semanas S1–S6
            if c0.startswith('Semana ') and c0 != 'Semana ':
                snum = c0.replace('Semana ','').strip()
                ali  = safe_float(row.iloc[1])
                beb  = safe_float(row.iloc[2])
                tot  = safe_float(row.iloc[3])
                pre  = safe_float(row.iloc[4])
                dif  = safe_float(row.iloc[5])
                pct  = safe_float(row.iloc[6])
                if presup_row is not None:
                    pre_cap = safe_float(df.iloc[presup_row, int(snum)])
                else:
                    pre_cap = pre
                if tot > 0 or pre_cap > 0:
                    d['semanas'].append({
                        'semana': f'Semana {snum}',
                        'alimentos': ali, 'bebidas': beb,
                        'total': tot, 'presupuesto': pre_cap,
                        'diferencia': tot - pre_cap,
                        'pct': (tot / pre_cap - 1) if pre_cap > 0 else 0
                    })

            elif c0 == 'TOTAL':
                d['total'] = {
                    'alimentos': safe_float(row.iloc[1]),
                    'bebidas':   safe_float(row.iloc[2]),
                    'total':     safe_float(row.iloc[3]),
                    'presupuesto': safe_float(row.iloc[4]),
                    'diferencia':  safe_float(row.iloc[5]),
                }

            elif c0 == 'Alimentos' and d['mix_ali'] is None:
                d['mix_ali'] = {
                    'total':    safe_float(row.iloc[1]),
                    'mix_real': safe_float(row.iloc[2]),
                    'meta':     safe_float(row.iloc[3]),
                }

            elif c0 == 'Bebidas' and d['mix_beb'] is None:
                d['mix_beb'] = {
                    'total':    safe_float(row.iloc[1]),
                    'mix_real': safe_float(row.iloc[2]),
                    'meta':     safe_float(row.iloc[3]),
                }

            elif c0 == 'Clientes':
                d['clientes'] = {
                    'real': safe_float(row.iloc[1]),
                    'meta': safe_float(row.iloc[2]),
                }

            elif c0 == 'Ticket Promedio':
                d['cheque'] = {
                    'real': safe_float(row.iloc[1]),
                    'meta': safe_float(row.iloc[2]),
                    'dif':  safe_float(row.iloc[3]),
                }
                in_resumen = False   # fin del bloque de este mes

        i += 1

    return meses_data

# ── Lectura staff hoja 2 ─────────────────────────────────────────────
def leer_staff(path):
    try:
        df = pd.read_excel(path, sheet_name=1, header=None)
        staff = []
        for i, row in df.iterrows():
            if i == 0: continue
            cant  = row.iloc[0]
            puest = row.iloc[1]
            if pd.notna(cant) and pd.notna(puest):
                staff.append({'cantidad': int(cant), 'puesto': str(puest).strip()})
        return staff
    except: return []

# ── Análisis por día de semana ────────────────────────────────────────
def analisis_por_dia(meses_data):
    """
    Agrupa todos los días registrados por día de semana (0=Lunes … 6=Domingo)
    y calcula promedios de venta total, ticket y comensales.
    Semana inicia lunes, termina domingo (según regla Argentilia).
    """
    from collections import defaultdict
    acum = defaultdict(lambda: {'ventas':[], 'tickets':[], 'comensales':[]})

    for mes, d in meses_data.items():
        for dia in d.get('dias', []):
            if dia['total'] > 0:
                dow = dia['dow']
                acum[dow]['ventas'].append(dia['total'])
                if dia['comensales'] > 0:
                    acum[dow]['tickets'].append(dia['total'] / dia['comensales'])
                    acum[dow]['comensales'].append(dia['comensales'])

    resultado = []
    for dow in range(7):
        v = acum[dow]
        resultado.append({
            'dia': DIAS_SEMANA[dow],
            'prom_venta':     round(sum(v['ventas']) / len(v['ventas']), 2) if v['ventas'] else 0,
            'prom_ticket':    round(sum(v['tickets']) / len(v['tickets']), 2) if v['tickets'] else 0,
            'prom_clientes':  round(sum(v['comensales']) / len(v['comensales']), 1) if v['comensales'] else 0,
            'n_dias':         len(v['ventas'])
        })
    return resultado

# ── Extracción completa ───────────────────────────────────────────────
def extraer_datos():
    datos = {}
    for nombre, archivo in ARCHIVOS.items():
        ruta = os.path.join(CARPETA, archivo)
        if not os.path.exists(ruta):
            print(f"⚠️  No encontré: {archivo}")
            continue
        print(f"✅ Leyendo: {nombre}")
        meses = leer_excel(ruta, nombre)
        datos[nombre] = {
            'meses': meses,
            'staff': leer_staff(ruta),
            'analisis_dia': analisis_por_dia(meses)
        }
    return datos

# ── Construcción JSON para el tablero ────────────────────────────────
def construir_js(datos):
    meses_con_data = set()
    for nombre, d in datos.items():
        for mes, md in d['meses'].items():
            if md.get('total') and md['total']['total'] > 0:
                meses_con_data.add(mes)
    meses_lista = [m for m in MESES_ES if m in meses_con_data]

    unidades = list(ARCHIVOS.keys())
    data_js  = {}

    for u in unidades:
        if u not in datos: continue
        meses = datos[u]['meses']

        def s(campo):
            return [meses.get(m, {}).get('total', {}).get(campo, 0) or 0 for m in meses_lista]

        def ticket(campo):
            return [meses.get(m, {}).get('cheque', {}).get(campo, 0) or 0 for m in meses_lista]

        def clientes(campo):
            return [meses.get(m, {}).get('clientes', {}).get(campo, 0) or 0 for m in meses_lista]

        def semanas():
            result = {}
            for mes in meses_lista:
                sems = meses.get(mes, {}).get('semanas', [])
                result[mes] = [{
                    's': s['semana'],
                    'a': round(s['alimentos']),
                    'b': round(s['bebidas']),
                    't': round(s['total']),
                    'p': round(s['presupuesto'])
                } for s in sems]
            return result

        data_js[u] = {
            'total':       s('total'),
            'presup':      s('presupuesto'),
            'alimentos':   s('alimentos'),
            'bebidas':     s('bebidas'),
            'ticket':      ticket('real'),
            'ticketMeta':  ticket('meta'),
            'clientes':    clientes('real'),
            'clientesMeta':clientes('meta'),
            'semanas':     semanas(),
            'staff':       datos[u]['staff'],
            'analisis_dia':datos[u]['analisis_dia'],
        }

    return meses_lista, data_js

# ── HTML ──────────────────────────────────────────────────────────────
def generar_html(meses, data, ultima_actualizacion):
    meses_json = json.dumps(meses, ensure_ascii=False)
    data_json  = json.dumps(data,  ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Tablero Ejecutivo — Grupo Gastronómico Argentilia</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  :root {{
    --red:#ED2E38;--gray-dark:#656266;--gray-mid:#B5B0AD;
    --gray-light:#F2F1F0;--white:#FFFFFF;--green:#1A7A4A;--amber:#D4860A;
  }}
  *{{margin:0;padding:0;box-sizing:border-box;}}
  body{{font-family:'Segoe UI',Arial,sans-serif;background:#EBEBEA;color:var(--gray-dark);}}
  .header{{background:var(--gray-dark);padding:20px 32px;display:flex;align-items:center;justify-content:space-between;}}
  .header-brand{{color:var(--red);font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;margin-bottom:4px;}}
  .header-title{{color:var(--white);font-size:20px;font-weight:700;}}
  .header-sub{{color:var(--gray-mid);font-size:12px;margin-top:2px;}}
  .header-badge{{background:rgba(237,46,56,0.15);border:1px solid var(--red);color:var(--red);padding:6px 14px;border-radius:4px;font-size:11px;font-weight:700;letter-spacing:1px;}}
  .timestamp{{background:rgba(237,46,56,0.08);border-left:3px solid var(--red);padding:8px 16px;font-size:11px;color:var(--gray-mid);}}
  .timestamp span{{color:var(--red);font-weight:700;}}
  .nav{{background:var(--white);border-bottom:2px solid var(--gray-light);padding:0 32px;display:flex;flex-wrap:wrap;}}
  .nav-btn{{padding:14px 18px;font-size:12px;font-weight:600;color:var(--gray-mid);border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;transition:all .2s;}}
  .nav-btn.active{{color:var(--red);border-bottom-color:var(--red);}}
  .content{{padding:24px 32px;}}
  .section{{display:none;}}.section.active{{display:block;}}
  .kpi-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px;}}
  .kpi-card{{background:var(--white);border-radius:8px;padding:18px 20px;border-left:4px solid var(--gray-mid);}}
  .kpi-card.positive{{border-left-color:var(--green);}}.kpi-card.negative{{border-left-color:var(--red);}}.kpi-card.neutral{{border-left-color:var(--amber);}}
  .kpi-label{{font-size:10px;font-weight:700;letter-spacing:1.5px;color:var(--gray-mid);text-transform:uppercase;margin-bottom:8px;}}
  .kpi-value{{font-size:22px;font-weight:700;color:var(--gray-dark);}}
  .kpi-sub{{font-size:11px;color:var(--gray-mid);margin-top:4px;}}
  .kpi-delta{{font-size:12px;font-weight:700;margin-top:4px;}}
  .kpi-delta.up{{color:var(--green);}}.kpi-delta.down{{color:var(--red);}}
  .chart-grid{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:24px;}}
  .chart-card{{background:var(--white);border-radius:8px;padding:20px;}}
  .chart-card.full{{grid-column:1/-1;}}
  .chart-title{{font-size:12px;font-weight:700;color:var(--gray-dark);letter-spacing:.5px;text-transform:uppercase;margin-bottom:16px;padding-bottom:10px;border-bottom:1px solid var(--gray-light);}}
  .chart-wrap{{position:relative;height:260px;}}
  .table-card{{background:var(--white);border-radius:8px;padding:20px;margin-bottom:20px;}}
  .table-title{{font-size:12px;font-weight:700;color:var(--gray-dark);letter-spacing:.5px;text-transform:uppercase;margin-bottom:16px;padding-bottom:10px;border-bottom:1px solid var(--gray-light);}}
  table{{width:100%;border-collapse:collapse;font-size:12.5px;}}
  th{{background:var(--gray-light);color:var(--gray-dark);font-weight:700;padding:9px 12px;text-align:left;font-size:11px;letter-spacing:.5px;}}
  td{{padding:9px 12px;border-bottom:1px solid var(--gray-light);}}
  tr:last-child td{{border-bottom:none;}}
  .badge{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;}}
  .badge-green{{background:#E6F4EC;color:var(--green);}}.badge-red{{background:#FDECEA;color:var(--red);}}.badge-amber{{background:#FEF3E2;color:var(--amber);}}
  .filter-bar{{display:flex;gap:8px;margin-bottom:20px;align-items:center;flex-wrap:wrap;}}
  .filter-label{{font-size:11px;font-weight:700;color:var(--gray-mid);letter-spacing:1px;text-transform:uppercase;margin-right:4px;}}
  .filter-btn{{padding:6px 14px;border-radius:20px;border:1px solid var(--gray-mid);background:var(--white);font-size:12px;color:var(--gray-dark);cursor:pointer;font-weight:600;transition:all .15s;}}
  .filter-btn.active{{background:var(--red);color:var(--white);border-color:var(--red);}}
  .rest-tabs{{display:flex;gap:8px;margin-bottom:20px;flex-wrap:wrap;}}
  .rest-tab{{padding:8px 18px;border-radius:6px;border:1px solid var(--gray-light);background:var(--white);font-size:12px;color:var(--gray-mid);cursor:pointer;font-weight:600;}}
  .rest-tab.active{{background:var(--gray-dark);color:var(--white);border-color:var(--gray-dark);}}
  .section-eyebrow{{font-size:10px;font-weight:700;letter-spacing:2px;color:var(--red);text-transform:uppercase;margin-bottom:6px;}}
  .section-heading{{font-size:16px;font-weight:700;color:var(--gray-dark);margin-bottom:20px;}}
  .note{{font-size:10px;color:var(--gray-mid);font-style:italic;margin-top:8px;}}
  .staff-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:12px;margin-top:4px;}}
  .staff-item{{background:var(--gray-light);border-radius:6px;padding:12px 16px;display:flex;align-items:center;gap:12px;}}
  .staff-num{{font-size:22px;font-weight:700;color:var(--red);min-width:32px;}}
  .staff-puesto{{font-size:12px;color:var(--gray-dark);font-weight:600;}}
  .dia-bar-wrap{{display:flex;align-items:center;gap:10px;margin-bottom:6px;}}
  .dia-label{{font-size:11px;color:var(--gray-dark);width:80px;font-weight:600;}}
  .dia-bar{{flex:1;height:18px;background:var(--gray-light);border-radius:4px;overflow:hidden;}}
  .dia-bar-fill{{height:18px;background:var(--red);border-radius:4px;transition:width .4s;}}
  .dia-val{{font-size:11px;color:var(--gray-dark);width:90px;text-align:right;}}
</style>
</head>
<body>
<div class="header">
  <div class="header-left">
    <div class="header-brand">Matus Consulting · Confidencial</div>
    <div class="header-title">Tablero Ejecutivo — Grupo Gastronómico Argentilia</div>
    <div class="header-sub">Héctor Vázquez · Director de Restaurantes Especializados</div>
  </div>
  <div class="header-badge">DIRECCIÓN OPERATIVA</div>
</div>
<div class="timestamp">Última actualización: <span>{ultima_actualizacion}</span></div>
<nav class="nav">
  <button class="nav-btn active" onclick="showSection('ranking',this)">Ranking & Comparativo</button>
  <button class="nav-btn" onclick="showSection('cumplimiento',this)">Cumplimiento Presupuestal</button>
  <button class="nav-btn" onclick="showSection('unidad',this)">Detalle por Unidad</button>
  <button class="nav-btn" onclick="showSection('mix',this)">Mix A&B · Ticket</button>
  <button class="nav-btn" onclick="showSection('dias',this)">Análisis por Día</button>
</nav>
<div class="content">

<!-- RANKING -->
<div id="ranking" class="section active">
  <div class="section-eyebrow">Visión Consolidada</div>
  <div class="section-heading">Ranking de Unidades · Acumulado del período</div>
  <div class="kpi-grid" id="kpi-ranking"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-title">Venta Total Mensual por Unidad (MXN)</div><div class="chart-wrap"><canvas id="chartComparativo"></canvas></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-title">Participación en Venta del Grupo</div><div class="chart-wrap"><canvas id="chartParticipacion"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">Mikoh — Evolución Mensual</div><div class="chart-wrap"><canvas id="chartMikoh"></canvas></div></div>
  </div>
  <div class="table-card">
    <div class="table-title">Ranking por Volumen de Venta</div>
    <table><thead><tr id="th-ranking"></tr></thead><tbody id="tabla-ranking"></tbody></table>
  </div>
</div>

<!-- CUMPLIMIENTO -->
<div id="cumplimiento" class="section">
  <div class="section-eyebrow">Análisis Presupuestal</div>
  <div class="section-heading">Cumplimiento vs Presupuesto · Por Unidad y Mes</div>
  <div class="filter-bar">
    <span class="filter-label">Unidad:</span>
    <button class="filter-btn active" onclick="setCumpFilter('todas',this)">Todas</button>
    <button class="filter-btn" onclick="setCumpFilter('Argentilia León',this)">A. León</button>
    <button class="filter-btn" onclick="setCumpFilter('Argentilia Querétaro',this)">A. Querétaro</button>
    <button class="filter-btn" onclick="setCumpFilter('Frascati',this)">Frascati</button>
    <button class="filter-btn" onclick="setCumpFilter('Mikoh',this)">Mikoh</button>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-title">% Cumplimiento Presupuestal Mensual</div><div class="chart-wrap"><canvas id="chartCumplimiento"></canvas></div></div>
  </div>
  <div class="table-card">
    <div class="table-title">Detalle de Cumplimiento</div>
    <table><thead><tr><th>Unidad</th><th>Mes</th><th>Venta Real</th><th>Presupuesto</th><th>Diferencia</th><th>% Cumpl.</th><th>Estatus</th></tr></thead>
    <tbody id="tabla-cumplimiento"></tbody></table>
  </div>
</div>

<!-- UNIDAD -->
<div id="unidad" class="section">
  <div class="section-eyebrow">Análisis Individual</div>
  <div class="section-heading">Desempeño Mensual por Unidad</div>
  <div class="rest-tabs" id="rest-tabs"></div>
  <div id="detalle-unidad"></div>
  <div class="table-card" style="margin-top:20px">
    <div class="table-title" id="tabla-semana-title">Cierre Semanal Detallado</div>
    <table><thead><tr><th>Mes</th><th>Semana</th><th>Alimentos</th><th>Bebidas</th><th>Total</th><th>Presupuesto</th><th>Diferencia</th><th>% vs Presup</th></tr></thead>
    <tbody id="tabla-semana-body"></tbody></table>
  </div>
</div>

<!-- MIX -->
<div id="mix" class="section">
  <div class="section-eyebrow">Análisis de Producto</div>
  <div class="section-heading">Mix Alimentos / Bebidas · Ticket Promedio</div>
  <div class="filter-bar" id="mix-filtros"><span class="filter-label">Mes:</span></div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-title">Mix A/B — <span id="mix-mes-label"></span></div><div class="chart-wrap"><canvas id="chartMixBar"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">Ticket Promedio vs Meta — <span id="mix-ticket-label"></span></div><div class="chart-wrap"><canvas id="chartTicketMes"></canvas></div></div>
  </div>
  <div class="table-card">
    <div class="table-title">Mix y Ticket por Unidad y Mes</div>
    <table><thead><tr><th>Unidad</th><th>Mes</th><th>Alimentos</th><th>% Alim</th><th>Bebidas</th><th>% Beb</th><th>Meta %Beb</th><th>Brecha</th><th>Ticket Real</th><th>Ticket Meta</th><th>Δ Ticket</th></tr></thead>
    <tbody id="tabla-mix"></tbody></table>
  </div>
</div>

<!-- ANÁLISIS POR DÍA -->
<div id="dias" class="section">
  <div class="section-eyebrow">Estrategia por Día</div>
  <div class="section-heading">Promedio de Venta, Ticket y Clientes por Día de Semana</div>
  <div class="rest-tabs" id="dias-rest-tabs"></div>
  <div class="chart-grid">
    <div class="chart-card">
      <div class="chart-title">Promedio de Venta por Día (MXN)</div>
      <div class="chart-wrap"><canvas id="chartDiaVenta"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Ticket Promedio por Día</div>
      <div class="chart-wrap"><canvas id="chartDiaTicket"></canvas></div>
    </div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full">
      <div class="chart-title">Promedio de Clientes por Día</div>
      <div class="chart-wrap"><canvas id="chartDiaClientes"></canvas></div>
    </div>
  </div>
  <div class="table-card">
    <div class="table-title" id="tabla-dias-title">Detalle por Día de Semana</div>
    <table>
      <thead><tr><th>Día</th><th>Prom. Venta</th><th>Prom. Ticket</th><th>Prom. Clientes</th><th>Días registrados</th><th>Índice vs Mejor día</th></tr></thead>
      <tbody id="tabla-dias-body"></tbody>
    </table>
  </div>
  <div class="table-card">
    <div class="table-title">Comparativo por Día — Todas las Unidades</div>
    <table><thead><tr><th>Día</th><th>A. León</th><th>A. Querétaro</th><th>Frascati</th><th>Mikoh</th><th>Mejor Unidad</th></tr></thead>
    <tbody id="tabla-dias-comp"></tbody></table>
  </div>
</div>

</div>

<script>
const MESES   = {meses_json};
const DATA    = {data_json};
const UNIDADES = Object.keys(DATA);
const COLORES  = {{'Argentilia León':'#656266','Argentilia Querétaro':'#ED2E38','Frascati':'#B5B0AD','Mikoh':'#1A7A4A'}};
const DIAS_SEMANA = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];
let charts = {{}};
let currentRest = UNIDADES[0];
let currentDiaRest = UNIDADES[0];
let currentCumpFilter = 'todas';
let currentMixMes = MESES[0];

function fmt(n){{ return '$'+Math.round(n||0).toLocaleString('es-MX'); }}
function fmtDec(n){{ return '$'+(n||0).toLocaleString('es-MX',{{minimumFractionDigits:2,maximumFractionDigits:2}}); }}
function pct(r,m){{ return m>0?((r-m)/m*100):null; }}

window.addEventListener('DOMContentLoaded',()=>{{
  buildKPIRanking(); buildTablaRanking(); buildComparativo(); buildParticipacion(); buildMikohTrend();
  buildCumplimiento(); fillTablaCumplimiento('todas');
  buildRestTabs(); buildDetalleUnidad(currentRest); fillTablaSemana(currentRest);
  buildMixFiltros(); buildMixBar(currentMixMes); buildTicketMes(currentMixMes); fillTablaMix();
  buildDiasRestTabs(); buildDiasCharts(currentDiaRest); fillTablaDias(currentDiaRest); fillTablaDiasComp();
}});

function showSection(id,btn){{
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
}}

// ── KPI Ranking ──────────────────────────────────────────────────────
function buildKPIRanking(){{
  const totales=UNIDADES.map(u=>DATA[u].total.reduce((a,b)=>a+b,0));
  const total_grupo=totales.reduce((a,b)=>a+b,0);
  const maxIdx=totales.indexOf(Math.max(...totales));
  const minTicket=UNIDADES.reduce((best,u)=>{{
    const t=DATA[u].ticket.filter(x=>x>0);
    const avg=t.length?t.reduce((a,b)=>a+b,0)/t.length:0;
    return avg<best.avg?{{u,avg}}:best;
  }},{{u:'',avg:99999}});
  const mk=DATA['Mikoh']?.total;
  const growth=mk&&mk.length>1?((mk[mk.length-1]-mk[0])/mk[0]*100).toFixed(1):0;
  document.getElementById('kpi-ranking').innerHTML=`
    <div class="kpi-card positive"><div class="kpi-label">Venta Total Acumulada</div><div class="kpi-value">${{fmt(total_grupo)}}</div><div class="kpi-sub">${{UNIDADES.length}} unidades · ${{MESES.length}} meses</div></div>
    <div class="kpi-card negative"><div class="kpi-label">#1 por Volumen</div><div class="kpi-value">${{UNIDADES[maxIdx].replace('Argentilia ','A. ')}}</div><div class="kpi-sub">${{fmt(totales[maxIdx])}}</div><div class="kpi-delta up">${{(totales[maxIdx]/total_grupo*100).toFixed(1)}}% del grupo</div></div>
    <div class="kpi-card neutral"><div class="kpi-label">Ticket más bajo</div><div class="kpi-value">${{minTicket.u.replace('Argentilia ','A. ')}}</div><div class="kpi-sub">Prom. $${{minTicket.avg.toFixed(0)}}</div></div>
    <div class="kpi-card positive"><div class="kpi-label">Mayor Crecimiento</div><div class="kpi-value">Mikoh +${{growth}}%</div><div class="kpi-sub">Primer vs último mes</div></div>`;
}}

function buildTablaRanking(){{
  const totales=UNIDADES.map(u=>DATA[u].total.reduce((a,b)=>a+b,0));
  const sorted=[...UNIDADES].sort((a,b)=>totales[UNIDADES.indexOf(b)]-totales[UNIDADES.indexOf(a)]);
  const medals=['🥇','🥈','🥉','4'];
  const colors=['#656266','#ED2E38','#656266','#B5B0AD'];
  document.getElementById('th-ranking').innerHTML='<th>Rank</th><th>Unidad</th>'+MESES.map(m=>`<th>${{m}}</th>`).join('')+'<th>Acumulado</th><th>Tendencia</th>';
  document.getElementById('tabla-ranking').innerHTML=sorted.map((u,i)=>{{
    const tu=totales[UNIDADES.indexOf(u)];
    const vals=DATA[u].total;
    const tend=vals.length>1?(vals[vals.length-1]>vals[0]?'↑ Crecimiento':vals[vals.length-1]<vals[0]?'↓ Caída':'→ Estable'):'—';
    const tc=tend.startsWith('↑')?'#1A7A4A':tend.startsWith('↓')?'#ED2E38':'#D4860A';
    return `<tr><td style="font-weight:700;color:#ED2E38">${{medals[i]}}</td><td style="font-weight:700;color:${{colors[i]}}">${{u}}</td>${{vals.map(v=>`<td>${{fmt(v)}}</td>`).join('')}}<td style="font-weight:700">${{fmt(tu)}}</td><td style="color:${{tc}};font-weight:700">${{tend}}</td></tr>`;
  }}).join('');
}}

function buildComparativo(){{
  if(charts.comp)charts.comp.destroy();
  charts.comp=new Chart(document.getElementById('chartComparativo').getContext('2d'),{{type:'bar',data:{{labels:MESES,datasets:UNIDADES.map(u=>({{"label":u.replace('Argentilia ','A. '),"data":DATA[u].total,"backgroundColor":COLORES[u]}}))}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}},scales:{{y:{{ticks:{{callback:v=>'$'+(v/1000000).toFixed(1)+'M',font:{{size:11}}}},grid:{{color:'#F2F1F0'}}}}}}}}}});
}}
function buildParticipacion(){{
  if(charts.part)charts.part.destroy();
  const totals=UNIDADES.map(u=>DATA[u].total.reduce((a,b)=>a+b,0));
  charts.part=new Chart(document.getElementById('chartParticipacion').getContext('2d'),{{type:'doughnut',data:{{labels:UNIDADES.map(u=>u.replace('Argentilia ','A. ')),datasets:[{{data:totals,backgroundColor:UNIDADES.map(u=>COLORES[u]),borderWidth:2,borderColor:'#fff'}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}}}}}});
}}
function buildMikohTrend(){{
  if(charts.mikoh)charts.mikoh.destroy();
  const d=DATA['Mikoh'];
  charts.mikoh=new Chart(document.getElementById('chartMikoh').getContext('2d'),{{type:'line',data:{{labels:MESES,datasets:[{{label:'Total',data:d.total,borderColor:'#1A7A4A',backgroundColor:'rgba(26,122,74,0.1)',fill:true,tension:0.3,pointRadius:5}},{{label:'Alimentos',data:d.alimentos,borderColor:'#656266',borderDash:[4,4],fill:false,tension:0.3,pointRadius:3}},{{label:'Bebidas',data:d.bebidas,borderColor:'#ED2E38',borderDash:[4,4],fill:false,tension:0.3,pointRadius:3}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}},scales:{{y:{{ticks:{{callback:v=>'$'+(v/1000000).toFixed(1)+'M',font:{{size:11}}}}}}}}}}}});
}}

// ── Cumplimiento ──────────────────────────────────────────────────────
function buildCumplimiento(){{
  if(charts.cump)charts.cump.destroy();
  const filtered=currentCumpFilter==='todas'?UNIDADES:[currentCumpFilter];
  charts.cump=new Chart(document.getElementById('chartCumplimiento').getContext('2d'),{{type:'line',data:{{labels:MESES,datasets:filtered.map(u=>({{"label":u.replace('Argentilia ','A. '),"data":MESES.map((m,i)=>{{const p=DATA[u].presup[i];return p>0?pct(DATA[u].total[i],p):null;}}),"borderColor":COLORES[u],"fill":false,"tension":0.2,"pointRadius":5}})) }},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}},scales:{{y:{{ticks:{{callback:v=>v+'%',font:{{size:11}}}},grid:{{color:'#F2F1F0'}}}}}}}}}});
}}
function setCumpFilter(f,btn){{
  currentCumpFilter=f;
  document.querySelectorAll('#cumplimiento .filter-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  buildCumplimiento();
  fillTablaCumplimiento(f);
}}
function fillTablaCumplimiento(f){{
  const units=f==='todas'?UNIDADES:[f];
  document.getElementById('tabla-cumplimiento').innerHTML=units.flatMap(u=>MESES.map((m,i)=>{{
    const r=DATA[u].total[i],p=DATA[u].presup[i];
    const pc=pct(r,p),dif=p>0?r-p:null;
    const badge=p===0?'<span class="badge" style="background:#F2F1F0;color:#B5B0AD">Sin meta</span>':pc>=0?'<span class="badge badge-green">✓ Alcanzado</span>':pc>=-10?'<span class="badge badge-amber">⚠ Brecha menor</span>':'<span class="badge badge-red">✗ Brecha crítica</span>';
    return `<tr><td><strong>${{u}}</strong></td><td>${{m}}</td><td>${{fmt(r)}}</td><td>${{p>0?fmt(p):'—'}}</td><td style="color:${{dif===null?'inherit':dif>=0?'#1A7A4A':'#ED2E38'}}">${{dif!==null?(dif>=0?'+':'')+fmt(dif):'—'}}</td><td style="font-weight:700;color:${{pc===null?'inherit':pc>=0?'#1A7A4A':'#ED2E38'}}">${{pc!==null?(pc>=0?'+':'')+pc.toFixed(1)+'%':'—'}}</td><td>${{badge}}</td></tr>`;
  }})).join('');
}}

// ── Detalle Unidad ─────────────────────────────────────────────────────
function buildRestTabs(){{
  document.getElementById('rest-tabs').innerHTML=UNIDADES.map((u,i)=>`<button class="rest-tab${{i===0?' active':''}}" onclick="selectRest('${{u}}',this)">${{u.replace('Argentilia ','A. ')}}</button>`).join('');
}}
function selectRest(u,btn){{
  currentRest=u;
  document.querySelectorAll('#unidad .rest-tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  buildDetalleUnidad(u);
  fillTablaSemana(u);
}}
function renderStaff(u){{
  const s=DATA[u]?.staff||[];
  if(!s.length)return '';
  return `<div class="table-card" style="margin-top:16px"><div class="table-title">Staff · ${{u}}</div><div class="staff-grid">${{s.map(r=>`<div class="staff-item"><div class="staff-num">${{r.cantidad}}</div><div class="staff-puesto">${{r.puesto}}</div></div>`).join('')}}</div></div>`;
}}
function buildDetalleUnidad(u){{
  const d=DATA[u];
  const mejor_idx=d.total.indexOf(Math.max(...d.total));
  const peor_idx=d.total.indexOf(Math.min(...d.total));
  const tickets=d.ticket.filter(x=>x>0);
  const avg_ticket=tickets.length?tickets.reduce((a,b)=>a+b,0)/tickets.length:0;
  const tend_ticket=tickets.length>1?(tickets[tickets.length-1]>tickets[0]?'↑ Subiendo':'↓ Cayendo'):'—';
  const tend_color=tend_ticket.startsWith('↑')?'#1A7A4A':'#ED2E38';
  const clientes_total=d.clientes.reduce((a,b)=>a+b,0);
  const kpis=`<div class="kpi-grid">
    <div class="kpi-card positive"><div class="kpi-label">Mejor Mes</div><div class="kpi-value">${{MESES[mejor_idx]}}</div><div class="kpi-sub">${{fmt(d.total[mejor_idx])}}</div></div>
    <div class="kpi-card negative"><div class="kpi-label">Mes más bajo</div><div class="kpi-value">${{MESES[peor_idx]}}</div><div class="kpi-sub">${{fmt(d.total[peor_idx])}}</div></div>
    <div class="kpi-card neutral"><div class="kpi-label">Ticket Promedio</div><div class="kpi-value" style="color:${{tend_color}}">${{tend_ticket}}</div><div class="kpi-sub">Prom. $${{avg_ticket.toFixed(0)}}</div></div>
    <div class="kpi-card positive"><div class="kpi-label">Comensales Período</div><div class="kpi-value">${{clientes_total.toLocaleString('es-MX')}}</div></div>
  </div>`;
  document.getElementById('detalle-unidad').innerHTML=kpis+renderStaff(u)+`<div class="chart-grid"><div class="chart-card full"><div class="chart-title">${{u}} · Alimentos y Bebidas vs Presupuesto</div><div class="chart-wrap"><canvas id="chartAB"></canvas></div></div></div>`;
  if(charts.ab)charts.ab.destroy();
  setTimeout(()=>{{
    charts.ab=new Chart(document.getElementById('chartAB').getContext('2d'),{{type:'bar',data:{{labels:MESES,datasets:[{{label:'Alimentos',data:d.alimentos,backgroundColor:'#656266',stack:'a'}},{{label:'Bebidas',data:d.bebidas,backgroundColor:'#ED2E38',stack:'a'}},{{label:'Presupuesto',data:d.presup.map(v=>v||null),backgroundColor:'rgba(0,0,0,0)',borderColor:'#B5B0AD',borderWidth:2,type:'line',pointRadius:4}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}},scales:{{y:{{stacked:true,ticks:{{callback:v=>'$'+(v/1000000).toFixed(1)+'M',font:{{size:11}}}}}},x:{{stacked:true}}}}}}}});
  }},50);
}}
function fillTablaSemana(u){{
  document.getElementById('tabla-semana-title').textContent=u+' · Cierre Semanal Detallado';
  document.getElementById('tabla-semana-body').innerHTML=MESES.flatMap(m=>{{
    const sems=(DATA[u].semanas||{{}})[m]||[];
    return sems.map(s=>{{
      const pc=s.p>0?((s.t-s.p)/s.p*100):null;
      const c=pc!==null?(pc>=0?'color:#1A7A4A':'color:#ED2E38'):'';
      return `<tr><td>${{m}}</td><td>${{s.s}}</td><td>${{fmt(s.a)}}</td><td>${{fmt(s.b)}}</td><td><strong>${{fmt(s.t)}}</strong></td><td>${{s.p>0?fmt(s.p):'—'}}</td><td style="${{c}}">${{pc!==null?(pc>=0?'+':'')+fmt(s.t-s.p):'—'}}</td><td style="${{c}};font-weight:700">${{pc!==null?(pc>=0?'+':'')+pc.toFixed(1)+'%':'—'}}</td></tr>`;
    }});
  }}).join('');
}}

// ── Mix ────────────────────────────────────────────────────────────────
function buildMixFiltros(){{
  const bar=document.getElementById('mix-filtros');
  MESES.forEach((m,i)=>{{
    const btn=document.createElement('button');
    btn.className='filter-btn'+(i===0?' active':'');
    btn.textContent=m;
    btn.onclick=()=>{{
      document.querySelectorAll('#mix-filtros .filter-btn').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
      currentMixMes=m;
      document.getElementById('mix-mes-label').textContent=m;
      document.getElementById('mix-ticket-label').textContent=m;
      buildMixBar(m);buildTicketMes(m);
    }};
    bar.appendChild(btn);
  }});
  document.getElementById('mix-mes-label').textContent=MESES[0];
  document.getElementById('mix-ticket-label').textContent=MESES[0];
}}
function buildMixBar(mes){{
  const idx=MESES.indexOf(mes);
  if(charts.mixBar)charts.mixBar.destroy();
  const labels=UNIDADES.map(u=>u.replace('Argentilia ','A. '));
  const ali=UNIDADES.map(u=>{{const t=DATA[u].alimentos[idx]+DATA[u].bebidas[idx];return t?(DATA[u].alimentos[idx]/t*100).toFixed(1):0;}});
  const beb=ali.map(a=>(100-parseFloat(a)).toFixed(1));
  charts.mixBar=new Chart(document.getElementById('chartMixBar').getContext('2d'),{{type:'bar',data:{{labels,datasets:[{{label:'Alimentos %',data:ali,backgroundColor:'#656266',stack:'mix'}},{{label:'Bebidas %',data:beb,backgroundColor:'#ED2E38',stack:'mix'}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}},scales:{{y:{{stacked:true,max:100,ticks:{{callback:v=>v+'%',font:{{size:11}}}}}},x:{{stacked:true}}}}}}}});
}}
function buildTicketMes(mes){{
  const idx=MESES.indexOf(mes);
  if(charts.ticketMes)charts.ticketMes.destroy();
  const labels=UNIDADES.map(u=>u.replace('Argentilia ','A. '));
  charts.ticketMes=new Chart(document.getElementById('chartTicketMes').getContext('2d'),{{type:'bar',data:{{labels,datasets:[{{label:'Ticket Real',data:UNIDADES.map(u=>DATA[u].ticket[idx]||0),backgroundColor:UNIDADES.map(u=>COLORES[u])}},{{label:'Meta',data:UNIDADES.map(u=>DATA[u].ticketMeta[idx]||null),backgroundColor:'rgba(0,0,0,0)',borderColor:'#656266',borderWidth:2,type:'line',pointRadius:5}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}},scales:{{y:{{ticks:{{callback:v=>'$'+v,font:{{size:11}}}}}}}}}}}});
}}
function fillTablaMix(){{
  document.getElementById('tabla-mix').innerHTML=UNIDADES.flatMap(u=>MESES.map((m,i)=>{{
    const a=DATA[u].alimentos[i],b=DATA[u].bebidas[i],t=a+b;
    const pA=t?(a/t*100).toFixed(1):'—',pB=t?(b/t*100).toFixed(1):'—';
    const brecha=t?(parseFloat(pB)-40).toFixed(1):null;
    const tk=DATA[u].ticket[i],tkM=DATA[u].ticketMeta[i];
    const dTk=tkM>0?(tk-tkM).toFixed(0):null;
    const cB=brecha!==null?(parseFloat(brecha)>=0?'color:#1A7A4A':'color:#ED2E38'):'';
    const cT=dTk!==null?(parseFloat(dTk)>=0?'color:#1A7A4A':'color:#ED2E38'):'';
    return `<tr><td><strong>${{u.replace('Argentilia ','A. ')}}</strong></td><td>${{m}}</td><td>${{fmt(a)}}</td><td>${{pA}}%</td><td>${{fmt(b)}}</td><td>${{pB}}%</td><td>40%</td><td style="${{cB}}">${{brecha!==null?(parseFloat(brecha)>=0?'+':'')+brecha+'pts':'—'}}</td><td>${{tk?fmtDec(tk):'—'}}</td><td>${{tkM>0?fmtDec(tkM):'—'}}</td><td style="${{cT}}">${{dTk!==null?(parseFloat(dTk)>=0?'+':'')+dTk:'—'}}</td></tr>`;
  }})).join('');
}}

// ── Análisis por Día ───────────────────────────────────────────────────
function buildDiasRestTabs(){{
  document.getElementById('dias-rest-tabs').innerHTML=UNIDADES.map((u,i)=>`<button class="rest-tab${{i===0?' active':''}}" onclick="selectDiaRest('${{u}}',this)">${{u.replace('Argentilia ','A. ')}}</button>`).join('');
}}
function selectDiaRest(u,btn){{
  currentDiaRest=u;
  document.querySelectorAll('#dias .rest-tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  buildDiasCharts(u);
  fillTablaDias(u);
}}
function buildDiasCharts(u){{
  const ad=DATA[u].analisis_dia||[];
  const labels=ad.map(d=>d.dia);
  const ventas=ad.map(d=>d.prom_venta);
  const tickets=ad.map(d=>d.prom_ticket);
  const clientes=ad.map(d=>d.prom_clientes);
  const color=COLORES[u];

  if(charts.diaVenta)charts.diaVenta.destroy();
  charts.diaVenta=new Chart(document.getElementById('chartDiaVenta').getContext('2d'),{{
    type:'bar',
    data:{{labels,datasets:[{{label:'Prom. Venta',data:ventas,backgroundColor:labels.map((l,i)=>ventas[i]===Math.max(...ventas)?'#ED2E38':color+'99')}}]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{y:{{ticks:{{callback:v=>'$'+(v/1000).toFixed(0)+'K',font:{{size:11}}}},grid:{{color:'#F2F1F0'}}}}}}}}
  }});

  if(charts.diaTicket)charts.diaTicket.destroy();
  charts.diaTicket=new Chart(document.getElementById('chartDiaTicket').getContext('2d'),{{
    type:'bar',
    data:{{labels,datasets:[{{label:'Prom. Ticket',data:tickets,backgroundColor:labels.map((l,i)=>tickets[i]===Math.max(...tickets)?'#1A7A4A':color+'99')}}]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{y:{{ticks:{{callback:v=>'$'+v.toFixed(0),font:{{size:11}}}},grid:{{color:'#F2F1F0'}}}}}}}}
  }});

  if(charts.diaClientes)charts.diaClientes.destroy();
  charts.diaClientes=new Chart(document.getElementById('chartDiaClientes').getContext('2d'),{{
    type:'bar',
    data:{{labels,datasets:[{{label:'Prom. Clientes',data:clientes,backgroundColor:labels.map((l,i)=>clientes[i]===Math.max(...clientes)?'#D4860A':color+'99')}}]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{y:{{ticks:{{font:{{size:11}}}},grid:{{color:'#F2F1F0'}}}}}}}}
  }});
}}
function fillTablaDias(u){{
  document.getElementById('tabla-dias-title').textContent=u+' · Detalle por Día de Semana';
  const ad=DATA[u].analisis_dia||[];
  const maxVenta=Math.max(...ad.map(d=>d.prom_venta));
  document.getElementById('tabla-dias-body').innerHTML=ad.map(d=>{{
    const idx=(d.prom_venta/maxVenta*100).toFixed(0);
    const c=d.prom_venta===maxVenta?'color:#ED2E38;font-weight:700':'';
    return `<tr>
      <td style="${{c}}">${{d.dia}}</td>
      <td style="${{c}}">${{fmt(d.prom_venta)}}</td>
      <td>${{fmtDec(d.prom_ticket)}}</td>
      <td>${{Math.round(d.prom_clientes)}}</td>
      <td style="color:#B5B0AD">${{d.n_dias}} días</td>
      <td>
        <div class="dia-bar-wrap">
          <div class="dia-bar"><div class="dia-bar-fill" style="width:${{idx}}%;background:${{d.prom_venta===maxVenta?'#ED2E38':'#B5B0AD'}}"></div></div>
          <div class="dia-val">${{idx}}%</div>
        </div>
      </td>
    </tr>`;
  }}).join('');
}}
function fillTablaDiasComp(){{
  const rows=DIAS_SEMANA.map((dia,dow)=>{{
    const vals=UNIDADES.map(u=>{{
      const ad=DATA[u].analisis_dia||[];
      return ad[dow]?ad[dow].prom_venta:0;
    }});
    const maxVal=Math.max(...vals);
    const bestU=UNIDADES[vals.indexOf(maxVal)];
    return `<tr><td><strong>${{dia}}</strong></td>
      ${{UNIDADES.map((u,i)=>`<td style="${{vals[i]===maxVal?'color:#1A7A4A;font-weight:700':''}}">${{fmt(vals[i])}}</td>`).join('')}}
      <td style="color:#1A7A4A;font-weight:700">${{bestU.replace('Argentilia ','A. ')}}</td>
    </tr>`;
  }}).join('');
  document.getElementById('tabla-dias-comp').innerHTML=rows;
}}
</script>
</body>
</html>"""
    return html

# ── MAIN ──────────────────────────────────────────────────────────────
print("\n🔄 Iniciando proceso de actualización...")
descargar_archivos()
print("\n📊 Leyendo archivos Excel...")
datos = extraer_datos()
print("\n⚙️  Construyendo datos para el tablero...")
meses, data = construir_js(datos)
print(f"   Meses detectados: {', '.join(meses)}")
print("\n🎨 Generando tablero HTML...")
html = generar_html(meses, data, ULTIMA_ACTUALIZACION)
salida = "index.html"
with open(salida, "w", encoding="utf-8") as f:
    f.write(html)
print(f"\n✅ ¡Listo! Tablero guardado como: {salida}")
print(f"   Última actualización: {ULTIMA_ACTUALIZACION}")
