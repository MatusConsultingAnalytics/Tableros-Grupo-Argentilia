import pandas as pd
import os
import json
import shutil
import requests
import calendar
from datetime import datetime, date, timedelta, timezone
try:
    from zoneinfo import ZoneInfo
    TZ_CDMX = ZoneInfo("America/Mexico_City")
except Exception:
    # Respaldo si el entorno no tiene la base tzdata disponible:
    # México central no usa horario de verano desde 2022 → UTC-6 fijo.
    TZ_CDMX = timezone(timedelta(hours=-6))
from io import BytesIO

_AHORA_CDMX = datetime.now(TZ_CDMX)
ULTIMA_ACTUALIZACION = _AHORA_CDMX.strftime("%d/%m/%Y %I:%M %p") + " (Hora CDMX)"

# ── Referencias de fecha para parcialidad y bandera de captura ────────
# HOY: fecha de corte (en horario de Ciudad de México) para saber qué
# semanas/días ya "se ejecutaron".
# FECHA_ESPERADA: último día que YA debería estar cargado (hoy - 1), ya que
# el día de hoy normalmente aún no cierra operación al momento de correr el script.
HOY = _AHORA_CDMX.date()
FECHA_ESPERADA = HOY - timedelta(days=1)

# ── Google Sheets (un solo archivo con varias hojas) ──────────────────
SHEETS_ID = "1gbO2DLTDFxN-kWseM7rExwawZ-Q5L8Pm"

HOJAS = {
    "Argentilia León":      "Argentilia León",
    "Argentilia Querétaro": "Argentilia Querétaro",
    "Frascati":             "Frascati",
    "Mikoh":                "Mikoh",
}

HOJAS_STAFF = {
    "Argentilia León":      "Argentilia León - Staff",
    "Argentilia Querétaro": "Argentilia Querétaro - Staff",
    "Frascati":             "Frascati - Staff",
    "Mikoh":                "Mikoh - Staff",
}

COMENSALES_SIMPLE = {"Argentilia León", "Mikoh"}

MESES_ORDEN = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
MESES_ES    = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
                "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

DIAS_SEMANA = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]

_xlsx_bytes = None

def descargar_archivos():
    global _xlsx_bytes
    url = f"https://docs.google.com/spreadsheets/d/{SHEETS_ID}/export?format=xlsx"
    print("⬇️  Descargando Google Sheets...")
    r = requests.get(url, timeout=60)
    if r.status_code != 200:
        raise Exception(f"Error al descargar Sheets: HTTP {r.status_code}")
    _xlsx_bytes = BytesIO(r.content)
    print(f"✅ Descarga completa ({len(r.content)//1024} KB)")

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

def agrupar_por_semana(dias):
    """Agrupa los días de un mes (ya leídos con fecha real) en semanas operativas
    Lunes–Domingo. Como cada mes es independiente, la primera y/o última semana
    puede quedar parcial (menos de 7 días) de forma natural — no se fuerza a 7."""
    dias_ordenados = sorted(dias, key=lambda d: d['fecha'])
    buckets, actual = [], []
    for d in dias_ordenados:
        if d['dow'] == 0 and actual:  # nuevo lunes → cierra la semana anterior
            buckets.append(actual)
            actual = []
        actual.append(d)
    if actual:
        buckets.append(actual)
    return buckets

def construir_calendario_mes(mes_nombre, dias_existentes):
    """Construye la lista COMPLETA de días calendario del mes (día 1 al último),
    sin depender de si la fila existe en el Excel. Si una fecha no tiene fila
    capturada, se agrega como 'sin dato' (total 0) — así el punto de partida
    para detectar atrasos es el calendario real, no lo que el archivo trae."""
    if mes_nombre not in MESES_ES:
        return dias_existentes
    mes_num = MESES_ES.index(mes_nombre) + 1
    anio = int(dias_existentes[0]['fecha'][:4]) if dias_existentes else HOY.year
    _, ultimo_dia = calendar.monthrange(anio, mes_num)
    por_fecha = {d['fecha']: d for d in dias_existentes}
    calendario = []
    for dia_num in range(1, ultimo_dia + 1):
        fecha_d = date(anio, mes_num, dia_num)
        fecha_str = fecha_d.strftime('%Y-%m-%d')
        if fecha_str in por_fecha:
            calendario.append(por_fecha[fecha_str])
        else:
            calendario.append({'fecha': fecha_str, 'dow': fecha_d.weekday(),
                                'alimentos': 0.0, 'bebidas': 0.0, 'total': 0.0, 'comensales': 0.0})
    return calendario

def enriquecer_parcialidad(mes_dict, mes_nombre):
    """Enriquece cada semana del resumen con: días que contiene, días ya
    transcurridos, si es futura/actual, y días sin captura. Con eso calcula:
    - objetivo_acumulado / venta_acumulada respetando la parcialidad (no suma
      semanas futuras y prorratea la semana en curso por días transcurridos).
    - bandera automática de atraso de captura para la semana en revisión.
    - tendencia proyectada para la semana operativa siguiente.
    """
    buckets = agrupar_por_semana(construir_calendario_mes(mes_nombre, mes_dict.get('dias', [])))
    semanas = mes_dict.get('semanas', [])

    objetivo_acum = 0.0
    venta_acum = 0.0

    for i, bucket in enumerate(buckets):
        if i >= len(semanas):
            continue
        sem = semanas[i]

        dias_en_semana = len(bucket)
        transcurridos = sum(1 for d in bucket if datetime.strptime(d['fecha'], '%Y-%m-%d').date() <= HOY)
        esperados, faltantes = 0, 0
        for d in bucket:
            fecha_d = datetime.strptime(d['fecha'], '%Y-%m-%d').date()
            if fecha_d <= FECHA_ESPERADA:
                esperados += 1
                if d['total'] == 0:
                    faltantes += 1

        sem['dias_en_semana']         = dias_en_semana
        sem['dias_transcurridos']     = transcurridos
        sem['es_futura']              = (transcurridos == 0)
        sem['es_actual']              = (0 < transcurridos < dias_en_semana)
        sem['dias_esperados_captura'] = esperados
        sem['dias_faltantes_captura'] = faltantes
        sem['fecha_inicio']           = bucket[0]['fecha']
        sem['fecha_fin']              = bucket[-1]['fecha']

        if not sem['es_futura']:
            venta_acum += sem['total']
            if sem['es_actual'] and dias_en_semana:
                objetivo_acum += sem['presupuesto'] * (transcurridos / dias_en_semana)
            else:
                objetivo_acum += sem['presupuesto']

    mes_dict['objetivo_acumulado']       = round(objetivo_acum, 2)
    mes_dict['venta_acumulada']          = round(venta_acum, 2)
    mes_dict['pct_cumplimiento_parcial'] = (venta_acum / objetivo_acum * 100) if objetivo_acum > 0 else None

    # ── Semana en revisión para la bandera ───────────────────────────
    # Es la semana que CONTIENE la fecha esperada (ayer) — no necesariamente
    # la semana en curso, ya que esta última puede llevar apenas 1 día y no
    # tener nada qué reportar todavía. Si el atraso viene de una semana ya
    # cerrada (ej. no se cargó nada desde hace 2 semanas), esa es la que
    # se debe señalar.
    semana_revision = None
    for sem in semanas[:len(buckets)]:
        fi = datetime.strptime(sem['fecha_inicio'], '%Y-%m-%d').date()
        ff = datetime.strptime(sem['fecha_fin'], '%Y-%m-%d').date()
        if fi <= FECHA_ESPERADA <= ff:
            semana_revision = sem
            break
    if semana_revision is None:
        candidatas = [s for s in semanas[:len(buckets)] if not s.get('es_futura', True)]
        semana_revision = candidatas[-1] if candidatas else None

    if semana_revision is not None:
        faltantes = semana_revision.get('dias_faltantes_captura', 0)
        nivel = 'rojo' if faltantes > 3 else ('amarillo' if faltantes >= 2 else 'ninguna')
        mes_dict['bandera'] = {
            'semana': semana_revision['semana'],
            'fecha_inicio': semana_revision['fecha_inicio'],
            'fecha_fin': semana_revision['fecha_fin'],
            'dias_faltantes': faltantes,
            'dias_esperados': semana_revision.get('dias_esperados_captura', 0),
            'nivel': nivel,
        }
    else:
        mes_dict['bandera'] = None

    # ── Tendencia a la semana operativa SIGUIENTE (no al mes) ──────────
    idx_actual = next((i for i, s in enumerate(semanas) if s.get('es_actual')), None)
    if idx_actual is None:
        cerradas = [i for i, s in enumerate(semanas) if not s.get('es_futura', True)]
        idx_actual = cerradas[-1] if cerradas else None

    tendencia = None
    if idx_actual is not None and idx_actual + 1 < len(semanas):
        siguiente = semanas[idx_actual + 1]
        cerradas_prev = [s for s in semanas[:idx_actual + 1]
                          if not s.get('es_futura') and not s.get('es_actual') and s.get('dias_en_semana')]
        ultimas2 = cerradas_prev[-2:]
        if ultimas2 and siguiente.get('presupuesto', 0) > 0:
            prom_diario = sum(s['total'] / s['dias_en_semana'] for s in ultimas2) / len(ultimas2)
            dias_siguiente = siguiente.get('dias_en_semana') or 7
            proyeccion = prom_diario * dias_siguiente
            tendencia = {
                'semana': siguiente['semana'],
                'proyeccion': round(proyeccion, 2),
                'presupuesto': siguiente['presupuesto'],
                'pct': round(proyeccion / siguiente['presupuesto'] * 100, 1),
            }
    mes_dict['tendencia_semana_siguiente'] = tendencia

def leer_excel(sheet_name, nombre):
    from io import BytesIO
    _xlsx_bytes.seek(0)
    df = pd.read_excel(_xlsx_bytes, sheet_name=sheet_name, header=None)
    simple = nombre in COMENSALES_SIMPLE
    meses_data  = {}
    current_mes = None
    presup_row  = None
    in_cap      = False
    in_resumen  = False

    i = 0
    while i < len(df):
        row  = df.iloc[i]
        c0   = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        c0up = c0.upper().replace(' ','')

        for idx, mk in enumerate(MESES_ORDEN):
            if c0up == mk:
                current_mes = MESES_ES[idx]
                meses_data[current_mes] = {
                    'semanas': [], 'total': None, 'cheque': None,
                    'clientes': None, 'mix_ali': None, 'mix_beb': None, 'dias': []
                }
                presup_row = None
                in_cap     = False
                in_resumen = False
                break

        if not current_mes:
            i += 1
            continue

        d = meses_data[current_mes]

        if c0 == 'Meta de venta ($)':
            presup_row = i
            i += 1
            continue

        if c0 == 'Fecha' and not in_resumen:
            in_cap = True
            i += 1
            continue

        if in_cap and not in_resumen:
            if c0 == 'RESUMEN AUTOMÁTICO (no editar — se calcula solo)':
                in_cap     = False
                in_resumen = True
                i += 1
                continue
            fecha = parse_fecha(row.iloc[0])
            if fecha:
                ali = safe_float(row.iloc[1])
                beb = safe_float(row.iloc[2])
                if simple:
                    com = safe_float(row.iloc[3])
                else:
                    com = safe_float(row.iloc[3]) + safe_float(row.iloc[4]) + safe_float(row.iloc[5])
                d['dias'].append({
                    'fecha': fecha.strftime('%Y-%m-%d'),
                    'dow': fecha.weekday(),
                    'alimentos': ali, 'bebidas': beb,
                    'total': ali + beb, 'comensales': com
                })
            i += 1
            continue

        if in_resumen:
            if c0.startswith('Semana ') and c0 != 'Semana ':
                snum = c0.replace('Semana ','').strip()
                try: snum_int = int(snum)
                except: i += 1; continue
                ali = safe_float(row.iloc[1])
                beb = safe_float(row.iloc[2])
                tot = safe_float(row.iloc[3])
                pre_cap = safe_float(df.iloc[presup_row, snum_int]) if presup_row is not None else safe_float(row.iloc[4])
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
                d['mix_ali'] = {'total': safe_float(row.iloc[1]), 'mix_real': safe_float(row.iloc[2]), 'meta': safe_float(row.iloc[3])}
            elif c0 == 'Bebidas' and d['mix_beb'] is None:
                d['mix_beb'] = {'total': safe_float(row.iloc[1]), 'mix_real': safe_float(row.iloc[2]), 'meta': safe_float(row.iloc[3])}
            elif c0 == 'Clientes':
                d['clientes'] = {'real': safe_float(row.iloc[1]), 'meta': safe_float(row.iloc[2])}
            elif c0 == 'Ticket Promedio':
                d['cheque'] = {'real': safe_float(row.iloc[1]), 'meta': safe_float(row.iloc[2]), 'dif': safe_float(row.iloc[3])}
                in_resumen = False
        i += 1

    for mes_nombre, mes_dict in meses_data.items():
        enriquecer_parcialidad(mes_dict, mes_nombre)

    return meses_data

def leer_staff(sheet_name):
    try:
        from io import BytesIO
        _xlsx_bytes.seek(0)
        df = pd.read_excel(_xlsx_bytes, sheet_name=sheet_name, header=None)
        staff = []
        for i, row in df.iterrows():
            if i == 0: continue
            cant = row.iloc[0]; puest = row.iloc[1]
            if pd.notna(cant) and pd.notna(puest):
                staff.append({'cantidad': int(cant), 'puesto': str(puest).strip()})
        return staff
    except: return []

def analisis_por_dia(meses_data, meses_filtro=None):
    from collections import defaultdict
    acum = defaultdict(lambda: {'ventas':[], 'tickets':[], 'comensales':[]})
    meses_iter = meses_filtro if meses_filtro else list(meses_data.keys())
    for mes in meses_iter:
        d = meses_data.get(mes, {})
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
            'prom_venta':    round(sum(v['ventas']) / len(v['ventas']), 2) if v['ventas'] else 0,
            'prom_ticket':   round(sum(v['tickets']) / len(v['tickets']), 2) if v['tickets'] else 0,
            'prom_clientes': round(sum(v['comensales']) / len(v['comensales']), 1) if v['comensales'] else 0,
            'n_dias': len(v['ventas'])
        })
    return resultado

def extraer_datos():
    datos = {}
    for nombre, sheet in HOJAS.items():
        print(f"✅ Leyendo: {nombre}")
        try:
            meses = leer_excel(sheet, nombre)
            staff = leer_staff(HOJAS_STAFF.get(nombre, ''))
            datos[nombre] = {'meses': meses, 'staff': staff}
        except Exception as e:
            print(f"⚠️  Error leyendo {nombre}: {e}")
    return datos

def construir_js(datos):
    meses_con_data = set()
    for nombre, d in datos.items():
        for mes, md in d['meses'].items():
            if md.get('total') and md['total']['total'] > 0:
                meses_con_data.add(mes)
    meses_lista = [m for m in MESES_ES if m in meses_con_data]

    data_js = {}
    for u in HOJAS.keys():
        if u not in datos: continue
        meses = datos[u]['meses']

        def s(campo): return [meses.get(m, {}).get('total', {}).get(campo, 0) or 0 for m in meses_lista]
        def ticket(c): return [meses.get(m, {}).get('cheque', {}).get(c, 0) or 0 for m in meses_lista]
        def clientes(c): return [meses.get(m, {}).get('clientes', {}).get(c, 0) or 0 for m in meses_lista]
        def semanas():
            r = {}
            for mes in meses_lista:
                sems = meses.get(mes, {}).get('semanas', [])
                r[mes] = [{'s': s['semana'], 'a': round(s['alimentos']), 'b': round(s['bebidas']),
                            't': round(s['total']), 'p': round(s['presupuesto']),
                            'dias': s.get('dias_en_semana'), 'transc': s.get('dias_transcurridos'),
                            'actual': s.get('es_actual', False), 'futura': s.get('es_futura', False),
                            'falt': s.get('dias_faltantes_captura', 0),
                            'fi': s.get('fecha_inicio'), 'ff': s.get('fecha_fin')} for s in sems]
            return r
        def dias_por_mes():
            r = {}
            for mes in meses_lista:
                r[mes] = meses.get(mes, {}).get('dias', [])
            return r
        def presup_parcial():
            return [meses.get(m, {}).get('objetivo_acumulado', 0) or 0 for m in meses_lista]
        def venta_parcial():
            return [meses.get(m, {}).get('venta_acumulada', 0) or 0 for m in meses_lista]
        def bandera_por_mes():
            return {m: meses.get(m, {}).get('bandera') for m in meses_lista}

        ultimo_mes = meses_lista[-1] if meses_lista else None
        tendencia_sem_sig = meses.get(ultimo_mes, {}).get('tendencia_semana_siguiente') if ultimo_mes else None

        data_js[u] = {
            'total': s('total'), 'presup': s('presupuesto'),
            'alimentos': s('alimentos'), 'bebidas': s('bebidas'),
            'ticket': ticket('real'), 'ticketMeta': ticket('meta'),
            'clientes': clientes('real'), 'clientesMeta': clientes('meta'),
            'semanas': semanas(),
            'dias_por_mes': dias_por_mes(),
            'staff': datos[u]['staff'],
            # ── Nuevo: parcialidad, bandera y tendencia ──────────────────
            'presupParcial': presup_parcial(),
            'ventaParcial': venta_parcial(),
            'banderaPorMes': bandera_por_mes(),
            'tendenciaSemSig': tendencia_sem_sig,
        }
    return meses_lista, data_js

def generar_html(meses, data, ultima_actualizacion):
    meses_json = json.dumps(meses, ensure_ascii=False)
    data_json  = json.dumps(data,  ensure_ascii=False)
    mes_actual_json = json.dumps(MESES_ES[HOY.month - 1], ensure_ascii=False)

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
  .timestamp{{background:rgba(237,46,56,0.08);border-left:3px solid var(--red);padding:8px 16px;font-size:11px;color:var(--gray-mid);display:flex;align-items:center;justify-content:space-between;gap:12px;}}
  .timestamp span{{color:var(--red);font-weight:700;}}
  .print-btn{{background:var(--white);border:1px solid var(--red);color:var(--red);padding:5px 14px;border-radius:4px;font-size:11px;font-weight:700;letter-spacing:.5px;cursor:pointer;white-space:nowrap;}}
  .print-btn:hover{{background:var(--red);color:var(--white);}}
  .nav{{background:var(--white);border-bottom:2px solid var(--gray-light);padding:0 32px;display:flex;flex-wrap:wrap;}}
  .nav-btn{{padding:14px 18px;font-size:12px;font-weight:600;color:var(--gray-mid);border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;transition:all .2s;}}
  .nav-btn.active{{color:var(--red);border-bottom-color:var(--red);}}
  .content{{padding:24px 32px;}}
  .section{{display:none;}}.section.active{{display:block;}}
  @media print{{
    body{{background:#fff;}}
    .nav, .print-btn{{display:none !important;}}
    .content{{padding:0 8px;}}
    .table-card, .chart-card, .kpi-card{{break-inside:avoid;}}
    .section{{display:none;}}
    .section.active{{display:block;}}
  }}
  .kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:24px;}}
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
  .filter-bar{{display:flex;gap:8px;margin-bottom:16px;align-items:center;flex-wrap:wrap;background:var(--white);padding:12px 16px;border-radius:8px;border:1px solid var(--gray-light);}}
  .filter-label{{font-size:10px;font-weight:700;color:var(--gray-mid);letter-spacing:1px;text-transform:uppercase;margin-right:4px;white-space:nowrap;}}
  .filter-btn{{padding:5px 12px;border-radius:20px;border:1px solid var(--gray-mid);background:var(--white);font-size:11px;color:var(--gray-dark);cursor:pointer;font-weight:600;transition:all .15s;white-space:nowrap;}}
  .filter-btn.active{{background:var(--red);color:var(--white);border-color:var(--red);}}
  .filter-divider{{width:1px;height:20px;background:var(--gray-light);margin:0 4px;}}
  .rest-tabs{{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap;}}
  .rest-tab{{padding:8px 18px;border-radius:6px;border:1px solid var(--gray-light);background:var(--white);font-size:12px;color:var(--gray-mid);cursor:pointer;font-weight:600;}}
  .rest-tab.active{{background:var(--gray-dark);color:var(--white);border-color:var(--gray-dark);}}
  .section-eyebrow{{font-size:10px;font-weight:700;letter-spacing:2px;color:var(--red);text-transform:uppercase;margin-bottom:6px;}}
  .section-heading{{font-size:16px;font-weight:700;color:var(--gray-dark);margin-bottom:16px;}}
  .periodo-badge{{display:inline-block;background:rgba(237,46,56,0.1);color:var(--red);border:1px solid var(--red);padding:3px 10px;border-radius:12px;font-size:11px;font-weight:700;margin-left:10px;vertical-align:middle;}}
  .note{{font-size:10px;color:var(--gray-mid);font-style:italic;margin-top:8px;}}
  .staff-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:12px;margin-top:4px;}}
  .staff-item{{background:var(--gray-light);border-radius:6px;padding:12px 16px;display:flex;align-items:center;gap:12px;}}
  .staff-num{{font-size:22px;font-weight:700;color:var(--red);min-width:32px;}}
  .staff-puesto{{font-size:12px;color:var(--gray-dark);font-weight:600;}}
  .dia-bar-wrap{{display:flex;align-items:center;gap:10px;margin-bottom:4px;}}
  .dia-label{{font-size:11px;color:var(--gray-dark);width:80px;font-weight:600;}}
  .dia-bar{{flex:1;height:16px;background:var(--gray-light);border-radius:4px;overflow:hidden;}}
  .dia-bar-fill{{height:16px;border-radius:4px;transition:width .4s;}}
  .dia-val{{font-size:11px;color:var(--gray-dark);width:90px;text-align:right;font-weight:600;}}
  .flag-banner{{padding:10px 16px;border-radius:6px;font-size:12.5px;font-weight:700;margin-bottom:16px;display:flex;align-items:center;gap:8px;}}
  .flag-banner.rojo{{background:#FDECEA;color:var(--red);border:1px solid var(--red);}}
  .flag-banner.amarillo{{background:#FEF3E2;color:var(--amber);border:1px solid var(--amber);}}
  .flag-banner.ninguna{{background:#E6F4EC;color:var(--green);border:1px solid var(--green);}}
  .trend-badge{{font-weight:700;}}
  .trend-badge.up{{color:var(--green);}}.trend-badge.flat{{color:var(--amber);}}.trend-badge.down{{color:var(--red);}}
  .methodology-note{{font-size:10.5px;color:var(--gray-mid);background:var(--white);border-left:3px solid var(--red);padding:8px 12px;margin-top:10px;line-height:1.5;}}
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
<div class="timestamp">
  <div>Última actualización: <span>{ultima_actualizacion}</span></div>
  <button class="print-btn" onclick="window.print()">🖨️ Imprimir esta sección</button>
</div>
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
    <div class="methodology-note">
      <strong>Objetivo Acumulado</strong>: suma de las semanas operativas ya concluidas + la semana en curso prorrateada por días transcurridos (no se suman semanas que aún no se ejecutan).
      <strong>Tendencia semana siguiente</strong>: proyección de venta para la próxima semana operativa (Lun–Dom), no para el mes, con base en el promedio diario de las últimas semanas cerradas vs. el objetivo de esa semana.
    </div>
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
    <table><thead><tr><th>Unidad</th><th>Mes</th><th>Venta Real (a la fecha)</th><th>Presupuesto Mensual (completo)</th><th>Objetivo a la fecha (parcial)</th><th>Diferencia</th><th>% Cumpl. (parcial)</th><th>Estatus</th></tr></thead>
    <tbody id="tabla-cumplimiento"></tbody></table>
    <div class="methodology-note">El <strong>% Cumpl. (parcial)</strong> compara la venta real contra el objetivo <em>a la fecha</em> (semanas ya ejecutadas + prorrateo de la semana en curso), no contra el presupuesto del mes completo. Así, un mes que aún no termina no se ve "en rojo" solo por no haber concluido.</div>
  </div>
</div>

<!-- UNIDAD -->
<div id="unidad" class="section">
  <div class="section-eyebrow">Análisis Individual</div>
  <div class="section-heading">Desempeño por Unidad <span class="periodo-badge" id="unidad-periodo-badge">Acumulado año</span></div>
  <div class="rest-tabs" id="rest-tabs"></div>
  <div class="filter-bar" id="unidad-periodo-bar"></div>
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

<!-- DÍAS -->
<div id="dias" class="section">
  <div class="section-eyebrow">Estrategia por Día</div>
  <div class="section-heading">Promedio por Día de Semana <span class="periodo-badge" id="dias-periodo-badge">Acumulado año</span></div>
  <div class="rest-tabs" id="dias-rest-tabs"></div>
  <div class="filter-bar" id="dias-periodo-bar"></div>
  <div id="dias-flag-banner"></div>
  <div class="kpi-grid" id="dias-kpis"></div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-title">Promedio de Venta por Día (MXN) <span style="font-weight:400;color:#B5B0AD;font-size:10px">— línea: objetivo diario promedio</span></div><div class="chart-wrap"><canvas id="chartDiaVenta"></canvas></div></div>
    <div class="chart-card"><div class="chart-title">Ticket Promedio por Día ($)</div><div class="chart-wrap"><canvas id="chartDiaTicket"></canvas></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-title">Promedio de Clientes por Día</div><div class="chart-wrap"><canvas id="chartDiaClientes"></canvas></div></div>
  </div>
  <div class="table-card">
    <div class="table-title" id="tabla-dias-title">Detalle por Día de Semana</div>
    <table><thead><tr><th>Día</th><th>Prom. Venta</th><th>Prom. Ticket</th><th>Prom. Clientes</th><th>Días registrados</th><th>Tendencia de Venta</th></tr></thead>
    <tbody id="tabla-dias-body"></tbody></table>
    <div class="methodology-note">El <strong>objetivo diario</strong> mostrado en la gráfica es el promedio del objetivo semanal entre sus días (no distingue día por día, ya que la meta se define por semana). Los promedios de venta ya excluyen días futuros sin captura, respetando la parcialidad del período seleccionado. La <strong>Tendencia de Venta</strong> compara el último registro real de ese día de la semana contra el promedio de sus registros previos dentro del período seleccionado (mínimo 2 registros para calcularse).</div>
  </div>
  <div class="table-card">
    <div class="table-title" id="tabla-dias-comp-title">Comparativo Semanal por Día — Evolución 4 Semanas</div>
    <div class="filter-bar" id="dias-comp-ancla-bar">
      <span class="filter-label">Semana de referencia:</span>
      <button class="filter-btn active" data-ancla="actual" onclick="setAnclaComparativo('actual',this)">Semana en curso (parcial)</button>
      <button class="filter-btn" data-ancla="pasada" onclick="setAnclaComparativo('pasada',this)">Última semana completa</button>
    </div>
    <table>
      <thead>
        <tr>
          <th rowspan="2">Día</th>
          <th colspan="4">Venta</th>
          <th colspan="4">Ticket Promedio</th>
          <th colspan="4">Clientes</th>
        </tr>
        <tr>
          <th id="th-comp-1">Sem. Actual</th><th>Sem. -1</th><th>Sem. -2</th><th>Sem. -3</th>
          <th id="th-comp-2">Sem. Actual</th><th>Sem. -1</th><th>Sem. -2</th><th>Sem. -3</th>
          <th id="th-comp-3">Sem. Actual</th><th>Sem. -1</th><th>Sem. -2</th><th>Sem. -3</th>
        </tr>
      </thead>
      <tbody id="tabla-dias-comp"></tbody>
    </table>
    <div class="methodology-note" id="tabla-dias-comp-note"></div>
  </div>
</div>

</div>

<script>
const MESES   = {meses_json};
const MES_ACTUAL = {mes_actual_json};
const DATA    = {data_json};
const UNIDADES = Object.keys(DATA);
const COLORES  = {{'Argentilia León':'#656266','Argentilia Querétaro':'#ED2E38','Frascati':'#B5B0AD','Mikoh':'#1A7A4A'}};
const DIAS_SEMANA = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'];

let charts = {{}};
let currentRest     = UNIDADES[0];
let currentDiaRest  = UNIDADES[0];
let currentCumpFilter = 'todas';
let currentMixMes   = MESES[0];

// ── Estado de período (compartido entre Detalle y Días) ────────────────
let periodoUnidad = 'año';
let periodoDias   = 'año';

function getMesesPeriodo(estado) {{
  const MESES_CERRADOS = MESES.filter(m => m !== MES_ACTUAL);
  if (estado === 'año')      return MESES_CERRADOS;
  if (estado === 'ultimos3') return MESES_CERRADOS.slice(-3);
  return [estado]; // mes específico, elegido explícitamente (puede ser el mes en curso)
}}

function etiquetaPeriodo(estado) {{
  if (estado === 'año')      return 'Acumulado año (meses cerrados)';
  if (estado === 'ultimos3') return 'Últimos 3 meses cerrados';
  return estado + (estado === MES_ACTUAL ? ' (en curso)' : '');
}}

// ── Calculadora de analisis_dia desde dias_por_mes ─────────────────────
function calcAnalisisdDia(u, mesesFiltro) {{
  const acum = Array.from({{length:7}}, ()=>({{'ventas':[],'tickets':[],'clientes':[],'serie':[]}}));
  mesesFiltro.forEach(mes => {{
    const dias = (DATA[u].dias_por_mes||{{}})[mes] || [];
    dias.forEach(d => {{
      if (d.total > 0) {{
        acum[d.dow].ventas.push(d.total);
        acum[d.dow].serie.push({{fecha:d.fecha, total:d.total}});
        if (d.comensales > 0) {{
          acum[d.dow].tickets.push(d.total / d.comensales);
          acum[d.dow].clientes.push(d.comensales);
        }}
      }}
    }});
  }});
  return acum.map((a,dow) => {{
    const serieOrdenada = a.serie.slice().sort((x,y)=> x.fecha<y.fecha?-1:(x.fecha>y.fecha?1:0));
    return {{
      dia: DIAS_SEMANA[dow],
      prom_venta:    a.ventas.length    ? a.ventas.reduce((x,y)=>x+y,0)/a.ventas.length       : 0,
      prom_ticket:   a.tickets.length   ? a.tickets.reduce((x,y)=>x+y,0)/a.tickets.length     : 0,
      prom_clientes: a.clientes.length  ? a.clientes.reduce((x,y)=>x+y,0)/a.clientes.length   : 0,
      n_dias: a.ventas.length,
      serie: serieOrdenada
    }};
  }});
}}

// ── Tendencia de venta por día de semana: último registro real vs. promedio de los anteriores ──
function tendenciaDia(serie){{
  if(!serie || serie.length < 2) return null;
  const n = serie.length;
  const ultimo = serie[n-1];
  const anteriores = serie.slice(0, n-1);
  const promAnt = anteriores.reduce((s,d)=>s+d.total,0)/anteriores.length;
  if(!promAnt) return null;
  return (ultimo.total-promAnt)/promAnt*100;
}}
function tendenciaBadgeHtml(pct){{
  if(pct===null) return '<span style="color:#B5B0AD;font-size:11px">Datos insuficientes</span>';
  const cls   = pct>1?'up':(pct<-1?'down':'flat');
  const arrow = pct>1?'▲':(pct<-1?'▼':'▬');
  const color = pct>1?'#1A7A4A':(pct<-1?'#ED2E38':'#B5B0AD');
  return `<span class="trend-badge ${{cls}}" style="color:${{color}}">${{arrow}} ${{pct>0?'+':''}}${{pct.toFixed(1)}}%</span>`;
}}

// ── Comparativo semanal por día: semana en curso vs. hace 3 semanas ────
function fmtFechaCorta(d){{
  const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'), day=String(d.getDate()).padStart(2,'0');
  return `${{y}}-${{m}}-${{day}}`;
}}
function parseFechaLocal(s){{ return new Date(s+'T00:00:00'); }}
function lunesDe(d){{
  const dd = new Date(d);
  const dow = dd.getDay(); // 0=Dom,1=Lun,...6=Sáb
  const diff = (dow===0?-6:1)-dow;
  dd.setDate(dd.getDate()+diff);
  dd.setHours(0,0,0,0);
  return dd;
}}
let anclaComparativoSemana = 'actual';
function setAnclaComparativo(val, btn){{
  anclaComparativoSemana = val;
  document.querySelectorAll('#dias-comp-ancla-bar .filter-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  fillTablaDiasComp(currentDiaRest);
}}
function buildComparativoSemanas(u, ancla){{
  const porMes = DATA[u].dias_por_mes||{{}};
  let flat=[];
  Object.values(porMes).forEach(arr=>{{ flat = flat.concat(arr); }});
  flat = flat.filter(d=>d.fecha && d.total>0);
  if(!flat.length) return null;
  const maxFecha = flat.reduce((max,d)=> d.fecha>max?d.fecha:max, flat[0].fecha);
  const lunesActual = lunesDe(parseFechaLocal(maxFecha));
  const mapSemana=(lunes)=>{{
    const dom = new Date(lunes); dom.setDate(dom.getDate()+6);
    const ini=fmtFechaCorta(lunes), fin=fmtFechaCorta(dom);
    const m={{}};
    flat.forEach(d=>{{ if(d.fecha>=ini && d.fecha<=fin) m[d.dow]=d; }});
    return {{mapa:m, ini, fin}};
  }};
  // El ancla define desde dónde se cuentan las 4 semanas: 'actual' = semana en
  // curso (offset 0, puede estar parcial); 'pasada' = última semana YA CERRADA
  // (offset 1), para poder comparar 7 días completos contra 7 días completos.
  const offsetBase = (ancla==='pasada') ? 1 : 0;
  const semanas = [0,1,2,3].map(i=>{{
    const offset = offsetBase + i;
    const lunes = new Date(lunesActual); lunes.setDate(lunes.getDate()-7*offset);
    return mapSemana(lunes);
  }});
  return {{ semanas }}; // semanas[0]=ancla, semanas[1]=ancla-1, semanas[2]=ancla-2, semanas[3]=ancla-3
}}
function fillTablaDiasComp(u){{
  const ancla = anclaComparativoSemana;
  const comp = buildComparativoSemanas(u, ancla);
  const titleEl = document.getElementById('tabla-dias-comp-title');
  const noteEl  = document.getElementById('tabla-dias-comp-note');
  const nombreU = u.replace('Argentilia ','A. ');
  const etiquetaAncla = ancla==='pasada' ? 'Última semana completa' : 'Semana en curso (parcial)';
  ['th-comp-1','th-comp-2','th-comp-3'].forEach(id=>{{ document.getElementById(id).textContent = ancla==='pasada' ? 'Sem. Pasada' : 'Sem. Actual'; }});
  if(!comp){{
    titleEl.textContent = `Comparativo Semanal por Día — ${{nombreU}}`;
    noteEl.textContent  = 'Sin información suficiente para construir el comparativo semanal.';
    document.getElementById('tabla-dias-comp').innerHTML = '';
    return;
  }}
  const [sAncla,s1,s2,s3] = comp.semanas;
  titleEl.textContent = `Comparativo Semanal por Día — ${{nombreU}} · ${{etiquetaAncla}} (${{sAncla.ini}} a ${{sAncla.fin}}) vs. Sem-1 (${{s1.ini}} a ${{s1.fin}}) · Sem-2 (${{s2.ini}} a ${{s2.fin}}) · Sem-3 (${{s3.ini}} a ${{s3.fin}})`;
  noteEl.innerHTML = ancla==='pasada'
    ? 'Se compara la <strong>última semana operativa ya cerrada</strong> (Lun–Dom completa) contra las 3 semanas operativas anteriores a esa — útil para un comparativo con los 7 días completos, sin huecos por días aún no transcurridos. Los días pueden pertenecer a meses distintos; eso no afecta el cálculo.'
    : 'Se compara la <strong>semana en curso</strong> (puede estar parcial si aún no termina) contra las 3 semanas operativas anteriores. Los días marcados "Pendiente" son días de la semana en curso que todavía no ocurren. Los días pueden pertenecer a meses distintos; eso no afecta el cálculo.';
  const delta=(a,p)=>(a===null||p===null||!p)?null:(a-p)/p*100;
  const celdaValor = (v, fmtFn) => v!==null ? fmtFn(v) : '—';
  const celdaComparativa = (actual, previo, fmtFn) => {{
    if(previo===null) return '<span style="color:#B5B0AD">—</span>';
    if(actual===null) return `<div>${{fmtFn(previo)}}</div><div style="margin-top:2px"><span style="color:#B5B0AD;font-size:11px">Pendiente</span></div>`;
    const d = delta(actual, previo);
    return `<div>${{fmtFn(previo)}}</div><div style="margin-top:2px">${{tendenciaBadgeHtml(d)}}</div>`;
  }};
  document.getElementById('tabla-dias-comp').innerHTML = DIAS_SEMANA.map((dia,dow)=>{{
    const dA=sAncla.mapa[dow], d1=s1.mapa[dow], d2=s2.mapa[dow], d3=s3.mapa[dow];
    const venta  = d => d?d.total:null;
    const ticket = d => (d&&d.comensales>0)?d.total/d.comensales:null;
    const cliente= d => d?d.comensales:null;
    const ventaA=venta(dA), tkA=ticket(dA), clA=cliente(dA);
    return `<tr>
      <td><strong>${{dia}}</strong></td>
      <td><strong>${{celdaValor(ventaA,fmt)}}</strong></td>
      <td>${{celdaComparativa(ventaA, venta(d1), fmt)}}</td>
      <td>${{celdaComparativa(ventaA, venta(d2), fmt)}}</td>
      <td>${{celdaComparativa(ventaA, venta(d3), fmt)}}</td>
      <td><strong>${{celdaValor(tkA,fmtDec)}}</strong></td>
      <td>${{celdaComparativa(tkA, ticket(d1), fmtDec)}}</td>
      <td>${{celdaComparativa(tkA, ticket(d2), fmtDec)}}</td>
      <td>${{celdaComparativa(tkA, ticket(d3), fmtDec)}}</td>
      <td><strong>${{celdaValor(clA,v=>Math.round(v).toString())}}</strong></td>
      <td>${{celdaComparativa(clA, cliente(d1), v=>Math.round(v).toString())}}</td>
      <td>${{celdaComparativa(clA, cliente(d2), v=>Math.round(v).toString())}}</td>
      <td>${{celdaComparativa(clA, cliente(d3), v=>Math.round(v).toString())}}</td>
    </tr>`;
  }}).join('');
}}

// ── Helpers ────────────────────────────────────────────────────────────
function fmt(n){{ return '$'+Math.round(n||0).toLocaleString('es-MX'); }}
function fmtDec(n){{ return '$'+(n||0).toLocaleString('es-MX',{{minimumFractionDigits:2,maximumFractionDigits:2}}); }}
function pct(r,m){{ return m>0?((r-m)/m*100):null; }}

window.addEventListener('DOMContentLoaded',()=>{{
  buildKPIRanking(); buildTablaRanking(); buildComparativo(); buildParticipacion(); buildMikohTrend();
  buildCumplimiento(); fillTablaCumplimiento('todas');
  buildRestTabs(); buildPeriodoBar('unidad-periodo-bar', ()=>periodoUnidad, v=>{{periodoUnidad=v;}}, refrescarUnidad, 'unidad-periodo-badge');
  buildDetalleUnidad(currentRest, getMesesPeriodo(periodoUnidad)); fillTablaSemana(currentRest, getMesesPeriodo(periodoUnidad));
  buildMixFiltros(); buildMixBar(currentMixMes); buildTicketMes(currentMixMes); fillTablaMix();
  buildDiasRestTabs(); buildPeriodoBar('dias-periodo-bar', ()=>periodoDias, v=>{{periodoDias=v;}}, refrescarDias, 'dias-periodo-badge');
  buildDiasCharts(currentDiaRest, getMesesPeriodo(periodoDias)); fillTablaDias(currentDiaRest, getMesesPeriodo(periodoDias)); fillTablaDiasComp(currentDiaRest);
  renderDiasExtras(currentDiaRest, getMesesPeriodo(periodoDias));
}});

function showSection(id,btn){{
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
}}

// ── Constructor de barra de período reutilizable ───────────────────────
function buildPeriodoBar(containerId, getEstado, setEstado, onRefresh, badgeId) {{
  const wrap = document.getElementById(containerId);
  if (!wrap) return;
  wrap.innerHTML = '';

  const addBtn = (val, label) => {{
    const btn = document.createElement('button');
    btn.className = 'filter-btn' + (getEstado() === val ? ' active' : '');
    btn.textContent = label;
    btn.onclick = () => {{
      setEstado(val);
      wrap.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      if (badgeId) document.getElementById(badgeId).textContent = etiquetaPeriodo(val);
      onRefresh();
    }};
    wrap.appendChild(btn);
  }};

  const lbl = document.createElement('span');
  lbl.className = 'filter-label';
  lbl.textContent = 'Período:';
  wrap.appendChild(lbl);

  addBtn('año', 'Acumulado año');
  addBtn('ultimos3', 'Últimos 3 meses');

  const div = document.createElement('span');
  div.className = 'filter-divider';
  wrap.appendChild(div);

  MESES.forEach(m => addBtn(m, m));
}}

function refrescarUnidad() {{
  const mp = getMesesPeriodo(periodoUnidad);
  buildDetalleUnidad(currentRest, mp);
  fillTablaSemana(currentRest, mp);
}}

function refrescarDias() {{
  const mp = getMesesPeriodo(periodoDias);
  buildDiasCharts(currentDiaRest, mp);
  fillTablaDias(currentDiaRest, mp);
  fillTablaDiasComp(currentDiaRest);
  renderDiasExtras(currentDiaRest, mp);
}}

// ── KPI Ranking ────────────────────────────────────────────────────────
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
  const objetivo_grupo=UNIDADES.reduce((a,u)=>a+(DATA[u].presupParcial||[]).reduce((x,y)=>x+y,0),0);
  const pct_grupo=objetivo_grupo>0?(total_grupo/objetivo_grupo*100):null;
  const pctColor=pct_grupo===null?'inherit':pct_grupo>=100?'#1A7A4A':pct_grupo>=90?'#D4860A':'#ED2E38';
  document.getElementById('kpi-ranking').innerHTML=`
    <div class="kpi-card positive"><div class="kpi-label">Venta Total Acumulada</div><div class="kpi-value">${{fmt(total_grupo)}}</div><div class="kpi-sub">${{UNIDADES.length}} unidades · ${{MESES.length}} meses</div></div>
    <div class="kpi-card neutral"><div class="kpi-label">Objetivo Acumulado (parcial)</div><div class="kpi-value">${{fmt(objetivo_grupo)}}</div><div class="kpi-sub">Respeta semanas ya ejecutadas</div></div>
    <div class="kpi-card ${{pct_grupo===null?'neutral':pct_grupo>=100?'positive':'negative'}}"><div class="kpi-label">% Cumplimiento Grupo</div><div class="kpi-value" style="color:${{pctColor}}">${{pct_grupo!==null?pct_grupo.toFixed(1)+'%':'—'}}</div><div class="kpi-sub">vs. objetivo a la fecha</div></div>
    <div class="kpi-card negative"><div class="kpi-label">#1 por Volumen</div><div class="kpi-value">${{UNIDADES[maxIdx].replace('Argentilia ','A. ')}}</div><div class="kpi-sub">${{fmt(totales[maxIdx])}}</div><div class="kpi-delta up">${{(totales[maxIdx]/total_grupo*100).toFixed(1)}}% del grupo</div></div>
    <div class="kpi-card neutral"><div class="kpi-label">Ticket más bajo</div><div class="kpi-value">${{minTicket.u.replace('Argentilia ','A. ')}}</div><div class="kpi-sub">Prom. $${{minTicket.avg.toFixed(0)}}</div></div>
    <div class="kpi-card positive"><div class="kpi-label">Mayor Crecimiento</div><div class="kpi-value">Mikoh +${{growth}}%</div><div class="kpi-sub">Primer vs último mes</div></div>`;
}}
function buildTablaRanking(){{
  const totales=UNIDADES.map(u=>DATA[u].total.reduce((a,b)=>a+b,0));
  const total_grupo=totales.reduce((a,b)=>a+b,0);
  const sorted=[...UNIDADES].sort((a,b)=>totales[UNIDADES.indexOf(b)]-totales[UNIDADES.indexOf(a)]);
  const medals=['🥇','🥈','🥉','4']; const colors=['#656266','#ED2E38','#656266','#B5B0AD'];
  document.getElementById('th-ranking').innerHTML='<th>Rank</th><th>Unidad</th>'+MESES.map(m=>`<th>${{m}}</th>`).join('')+'<th>Venta Acumulada</th><th>Objetivo Acumulado</th><th>% Cumplimiento</th><th>% Participación</th><th>Tendencia Sem. Siguiente</th>';
  document.getElementById('tabla-ranking').innerHTML=sorted.map((u,i)=>{{
    const tu=totales[UNIDADES.indexOf(u)]; const vals=DATA[u].total;
    const objetivoAcum=(DATA[u].presupParcial||[]).reduce((a,b)=>a+b,0);
    const pctCump=objetivoAcum>0?(tu/objetivoAcum*100):null;
    const pctColor=pctCump===null?'inherit':pctCump>=100?'#1A7A4A':pctCump>=90?'#D4860A':'#ED2E38';
    const participacion=total_grupo>0?(tu/total_grupo*100).toFixed(1):'0.0';
    const ts=DATA[u].tendenciaSemSig;
    let tendHtml='<span style="color:#B5B0AD">Sin datos suficientes</span>';
    if(ts){{
      const cls=ts.pct>=100?'up':ts.pct>=90?'flat':'down';
      const icon=ts.pct>=100?'▲':ts.pct>=90?'▶':'▼';
      tendHtml=`<span class="trend-badge ${{cls}}">${{icon}} ${{ts.pct.toFixed(0)}}%</span> <span style="color:#B5B0AD;font-size:10px">(${{ts.semana}})</span>`;
    }}
    return `<tr><td style="font-weight:700;color:#ED2E38">${{medals[i]}}</td><td style="font-weight:700;color:${{colors[i]}}">${{u}}</td>${{vals.map(v=>`<td>${{fmt(v)}}</td>`).join('')}}<td style="font-weight:700">${{fmt(tu)}}</td><td>${{fmt(objetivoAcum)}}</td><td style="font-weight:700;color:${{pctColor}}">${{pctCump!==null?pctCump.toFixed(1)+'%':'—'}}</td><td>${{participacion}}%</td><td>${{tendHtml}}</td></tr>`;
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

// ── Cumplimiento ────────────────────────────────────────────────────────
function buildCumplimiento(){{
  if(charts.cump)charts.cump.destroy();
  const filtered=currentCumpFilter==='todas'?UNIDADES:[currentCumpFilter];
  charts.cump=new Chart(document.getElementById('chartCumplimiento').getContext('2d'),{{type:'line',data:{{labels:MESES,datasets:filtered.map(u=>({{"label":u.replace('Argentilia ','A. '),"data":MESES.map((m,i)=>{{const p=(DATA[u].presupParcial||[])[i];return p>0?pct(DATA[u].total[i],p):null;}}),"borderColor":COLORES[u],"fill":false,"tension":0.2,"pointRadius":5}})) }},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}},scales:{{y:{{ticks:{{callback:v=>v+'%',font:{{size:11}}}},grid:{{color:'#F2F1F0'}}}}}}}}}});
}}
function setCumpFilter(f,btn){{
  currentCumpFilter=f;
  document.querySelectorAll('#cumplimiento .filter-bar .filter-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); buildCumplimiento(); fillTablaCumplimiento(f);
}}
function fillTablaCumplimiento(f){{
  const units=f==='todas'?UNIDADES:[f];
  document.getElementById('tabla-cumplimiento').innerHTML=units.flatMap(u=>MESES.map((m,i)=>{{
    const r=DATA[u].total[i],pMensual=DATA[u].presup[i];
    const pParcial=(DATA[u].presupParcial||[])[i]||0;
    const pc=pct(r,pParcial),dif=pParcial>0?r-pParcial:null;
    const badge=pParcial===0?'<span class="badge" style="background:#F2F1F0;color:#B5B0AD">Sin meta</span>':pc>=0?'<span class="badge badge-green">✓ Alcanzado</span>':pc>=-10?'<span class="badge badge-amber">⚠ Brecha menor</span>':'<span class="badge badge-red">✗ Brecha crítica</span>';
    return `<tr><td><strong>${{u}}</strong></td><td>${{m}}</td><td>${{fmt(r)}}</td><td>${{pMensual>0?fmt(pMensual):'—'}}</td><td>${{pParcial>0?fmt(pParcial):'—'}}</td><td style="color:${{dif===null?'inherit':dif>=0?'#1A7A4A':'#ED2E38'}}">${{dif!==null?(dif>=0?'+':'')+fmt(dif):'—'}}</td><td style="font-weight:700;color:${{pc===null?'inherit':pc>=0?'#1A7A4A':'#ED2E38'}}">${{pc!==null?(pc>=0?'+':'')+pc.toFixed(1)+'%':'—'}}</td><td>${{badge}}</td></tr>`;
  }})).join('');
}}

// ── Detalle por Unidad ──────────────────────────────────────────────────
function buildRestTabs(){{
  document.getElementById('rest-tabs').innerHTML=UNIDADES.map((u,i)=>`<button class="rest-tab${{i===0?' active':''}}" onclick="selectRest('${{u}}',this)">${{u.replace('Argentilia ','A. ')}}</button>`).join('');
}}
function selectRest(u,btn){{
  currentRest=u;
  document.querySelectorAll('#unidad .rest-tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const mp=getMesesPeriodo(periodoUnidad);
  buildDetalleUnidad(u,mp); fillTablaSemana(u,mp);
}}
function renderStaff(u){{
  const s=DATA[u]?.staff||[];
  if(!s.length)return '';
  return `<div class="table-card" style="margin-top:16px"><div class="table-title">Staff · ${{u}}</div><div class="staff-grid">${{s.map(r=>`<div class="staff-item"><div class="staff-num">${{r.cantidad}}</div><div class="staff-puesto">${{r.puesto}}</div></div>`).join('')}}</div></div>`;
}}
function buildDetalleUnidad(u, mesesFiltro){{
  const d=DATA[u];
  const idxList=mesesFiltro.map(m=>MESES.indexOf(m)).filter(i=>i>=0);
  const totFiltro=idxList.map(i=>d.total[i]||0);
  const mejor_idx=idxList[totFiltro.indexOf(Math.max(...totFiltro))];
  const peor_idx=idxList[totFiltro.indexOf(Math.min(...totFiltro))];
  const tickets=idxList.map(i=>d.ticket[i]).filter(x=>x>0);
  const avg_ticket=tickets.length?tickets.reduce((a,b)=>a+b,0)/tickets.length:0;
  const tend_ticket=tickets.length>1?(tickets[tickets.length-1]>tickets[0]?'↑ Subiendo':'↓ Cayendo'):'—';
  const tend_color=tend_ticket.startsWith('↑')?'#1A7A4A':'#ED2E38';
  const clientes_total=idxList.reduce((s,i)=>s+(d.clientes[i]||0),0);
  const venta_total=idxList.reduce((s,i)=>s+(d.total[i]||0),0);
  const objetivo_total=idxList.reduce((s,i)=>s+((d.presupParcial||[])[i]||0),0);
  const pctCump=objetivo_total>0?(venta_total/objetivo_total*100):null;
  let evalBadge='<span class="badge" style="background:#F2F1F0;color:#B5B0AD">Sin meta</span>';
  if(pctCump!==null){{
    evalBadge=pctCump>=100?'<span class="badge badge-green">✓ Cumplido</span>':pctCump>=90?'<span class="badge badge-amber">⚠ En riesgo</span>':'<span class="badge badge-red">✗ No cumplido</span>';
  }}
  const ts=d.tendenciaSemSig;
  let tendHtml='—', tendSub='Sin datos suficientes';
  if(ts){{
    const cls=ts.pct>=100?'up':ts.pct>=90?'flat':'down';
    const icon=ts.pct>=100?'▲':ts.pct>=90?'▶':'▼';
    tendHtml=`<span class="trend-badge ${{cls}}">${{icon}} ${{ts.pct.toFixed(0)}}%</span>`;
    tendSub=`Proyección ${{ts.semana}}`;
  }}
  const kpis=`<div class="kpi-grid">
    <div class="kpi-card positive"><div class="kpi-label">Venta del Período</div><div class="kpi-value">${{fmt(venta_total)}}</div><div class="kpi-sub">${{mesesFiltro.join(' · ')}}</div></div>
    <div class="kpi-card neutral"><div class="kpi-label">Objetivo Acumulado</div><div class="kpi-value">${{fmt(objetivo_total)}}</div><div class="kpi-sub">Respeta semanas ya ejecutadas</div></div>
    <div class="kpi-card ${{pctCump===null?'neutral':pctCump>=100?'positive':'negative'}}"><div class="kpi-label">% Cumplimiento</div><div class="kpi-value">${{pctCump!==null?pctCump.toFixed(1)+'%':'—'}}</div><div class="kpi-sub">${{evalBadge}}</div></div>
    <div class="kpi-card neutral"><div class="kpi-label">Tendencia Sem. Siguiente</div><div class="kpi-value">${{tendHtml}}</div><div class="kpi-sub">${{tendSub}}</div></div>
    <div class="kpi-card positive"><div class="kpi-label">Mejor Mes</div><div class="kpi-value">${{mejor_idx>=0?MESES[mejor_idx]:'—'}}</div><div class="kpi-sub">${{mejor_idx>=0?fmt(d.total[mejor_idx]):'—'}}</div></div>
    <div class="kpi-card neutral"><div class="kpi-label">Ticket Promedio</div><div class="kpi-value" style="color:${{tend_color}}">${{tend_ticket}}</div><div class="kpi-sub">Prom. $${{avg_ticket.toFixed(0)}}</div></div>
    <div class="kpi-card positive"><div class="kpi-label">Comensales</div><div class="kpi-value">${{clientes_total.toLocaleString('es-MX')}}</div><div class="kpi-sub">En el período</div></div>
  </div>`;
  document.getElementById('detalle-unidad').innerHTML=kpis+renderStaff(u)+`<div class="chart-grid"><div class="chart-card full"><div class="chart-title">${{u}} · Alimentos y Bebidas vs Presupuesto</div><div class="chart-wrap"><canvas id="chartAB"></canvas></div></div></div>`;
  if(charts.ab)charts.ab.destroy();
  setTimeout(()=>{{
    const labels=mesesFiltro;
    const ali=idxList.map(i=>d.alimentos[i]||0);
    const beb=idxList.map(i=>d.bebidas[i]||0);
    const pre=idxList.map(i=>d.presup[i]||null);
    const pctAli=idxList.map(i=>{{const t=(d.alimentos[i]||0)+(d.bebidas[i]||0); return t?+((d.alimentos[i]/t)*100).toFixed(1):null;}});
    charts.ab=new Chart(document.getElementById('chartAB').getContext('2d'),{{type:'bar',data:{{labels,datasets:[
      {{label:'Alimentos',data:ali,backgroundColor:'#656266',stack:'a'}},
      {{label:'Bebidas',data:beb,backgroundColor:'#ED2E38',stack:'a'}},
      {{label:'Presupuesto',data:pre,backgroundColor:'rgba(0,0,0,0)',borderColor:'#B5B0AD',borderWidth:2,type:'line',pointRadius:4}},
      {{label:'% Alimentos (mezcla)',data:pctAli,type:'line',yAxisID:'y1',borderColor:'#1A7A4A',backgroundColor:'#1A7A4A',borderDash:[3,3],tension:.25,pointRadius:3}}
    ]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}},scales:{{
      y:{{stacked:true,ticks:{{callback:v=>'$'+(v/1000000).toFixed(1)+'M',font:{{size:11}}}}}},
      y1:{{beginAtZero:true,max:100,position:'right',grid:{{drawOnChartArea:false}},ticks:{{callback:v=>v+'%',font:{{size:11}}}}}},
      x:{{stacked:true}}}}}}}});
  }},50);
}}
function fillTablaSemana(u, mesesFiltro){{
  document.getElementById('tabla-semana-title').textContent=u+' · Cierre Semanal Detallado';
  document.getElementById('tabla-semana-body').innerHTML=mesesFiltro.flatMap(m=>{{
    const sems=(DATA[u].semanas||{{}})[m]||[];
    return sems.map(s=>{{
      const pc=s.p>0?((s.t-s.p)/s.p*100):null;
      const c=pc!==null?(pc>=0?'color:#1A7A4A':'color:#ED2E38'):'';
      return `<tr><td>${{m}}</td><td>${{s.s}}</td><td>${{fmt(s.a)}}</td><td>${{fmt(s.b)}}</td><td><strong>${{fmt(s.t)}}</strong></td><td>${{s.p>0?fmt(s.p):'—'}}</td><td style="${{c}}">${{pc!==null?(pc>=0?'+':'')+fmt(s.t-s.p):'—'}}</td><td style="${{c}};font-weight:700">${{pc!==null?(pc>=0?'+':'')+pc.toFixed(1)+'%':'—'}}</td></tr>`;
    }});
  }}).join('');
}}

// ── Mix ─────────────────────────────────────────────────────────────────
function buildMixFiltros(){{
  const bar=document.getElementById('mix-filtros');
  MESES.forEach((m,i)=>{{
    const btn=document.createElement('button');
    btn.className='filter-btn'+(i===0?' active':'');
    btn.textContent=m;
    btn.onclick=()=>{{
      document.querySelectorAll('#mix-filtros .filter-btn').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active'); currentMixMes=m;
      document.getElementById('mix-mes-label').textContent=m;
      document.getElementById('mix-ticket-label').textContent=m;
      buildMixBar(m); buildTicketMes(m);
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

// ── Análisis por Día ────────────────────────────────────────────────────
function buildDiasRestTabs(){{
  document.getElementById('dias-rest-tabs').innerHTML=UNIDADES.map((u,i)=>`<button class="rest-tab${{i===0?' active':''}}" onclick="selectDiaRest('${{u}}',this)">${{u.replace('Argentilia ','A. ')}}</button>`).join('');
}}
function selectDiaRest(u,btn){{
  currentDiaRest=u;
  document.querySelectorAll('#dias .rest-tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const mp=getMesesPeriodo(periodoDias);
  buildDiasCharts(u,mp); fillTablaDias(u,mp); fillTablaDiasComp(u); renderDiasExtras(u,mp);
}}

// ── Bandera automática de captura + KPIs de objetivo/acumulado ─────────
function objetivoDiarioPromedio(u, mesesFiltro){{
  let ps=[];
  mesesFiltro.forEach(m=>{{
    (((DATA[u].semanas||{{}})[m])||[]).forEach(s=>{{
      if(!s.futura && s.dias) ps.push(s.p/s.dias);
    }});
  }});
  return ps.length? ps.reduce((a,b)=>a+b,0)/ps.length : 0;
}}
function renderDiasExtras(u, mesesFiltro){{
  // KPIs del período (respetan la parcialidad, igual que Ranking/Cumplimiento)
  const idxList=mesesFiltro.map(m=>MESES.indexOf(m)).filter(i=>i>=0);
  const venta=idxList.reduce((s,i)=>s+(DATA[u].total[i]||0),0);
  const objetivo=idxList.reduce((s,i)=>s+((DATA[u].presupParcial||[])[i]||0),0);
  const pc=objetivo>0?(venta/objetivo*100):null;
  document.getElementById('dias-kpis').innerHTML=`
    <div class="kpi-card positive"><div class="kpi-label">Venta Acumulada del Período</div><div class="kpi-value">${{fmt(venta)}}</div></div>
    <div class="kpi-card neutral"><div class="kpi-label">Objetivo Acumulado (parcial)</div><div class="kpi-value">${{fmt(objetivo)}}</div></div>
    <div class="kpi-card ${{pc===null?'neutral':pc>=100?'positive':'negative'}}"><div class="kpi-label">% Cumplimiento</div><div class="kpi-value">${{pc!==null?pc.toFixed(1)+'%':'—'}}</div></div>`;

  // Bandera: se evalúa sobre el último mes del período seleccionado que tenga bandera calculada
  const banner=document.getElementById('dias-flag-banner');
  let mesConBandera=null, band=null;
  for(let i=mesesFiltro.length-1;i>=0;i--){{
    const b=(DATA[u].banderaPorMes||{{}})[mesesFiltro[i]];
    if(b){{ mesConBandera=mesesFiltro[i]; band=b; break; }}
  }}
  if(!band){{ banner.innerHTML=''; return; }}
  if(band.nivel==='rojo'){{
    banner.innerHTML=`<div class="flag-banner rojo">🔴 Sin información disponible para análisis — ${{band.semana}} (${{band.fecha_inicio}} a ${{band.fecha_fin}}): ${{band.dias_faltantes}} días sin captura.</div>`;
  }} else if(band.nivel==='amarillo'){{
    banner.innerHTML=`<div class="flag-banner amarillo">🟡 Captura incompleta — ${{band.semana}} (${{band.fecha_inicio}} a ${{band.fecha_fin}}): ${{band.dias_faltantes}} días sin captura. El análisis debe tomarse con reserva.</div>`;
  }} else {{
    banner.innerHTML=`<div class="flag-banner ninguna">🟢 Información al corriente — captura completa hasta el día anterior a la actualización.</div>`;
  }}
}}
function buildDiasCharts(u, mesesFiltro){{
  const ad=calcAnalisisdDia(u, mesesFiltro);
  const labels=ad.map(d=>d.dia);
  const ventas=ad.map(d=>d.prom_venta);
  const tickets=ad.map(d=>d.prom_ticket);
  const clientes=ad.map(d=>d.prom_clientes);
  const color=COLORES[u];
  const maxV=Math.max(...ventas), maxT=Math.max(...tickets), maxC=Math.max(...clientes);

  const objDiario=objetivoDiarioPromedio(u, mesesFiltro);

  if(charts.diaVenta)charts.diaVenta.destroy();
  charts.diaVenta=new Chart(document.getElementById('chartDiaVenta').getContext('2d'),{{
    type:'bar',
    data:{{labels,datasets:[
      {{label:'Prom. Venta',data:ventas,backgroundColor:ventas.map(v=>v===maxV?'#ED2E38':color+'88')}},
      {{label:'Objetivo diario (prom.)',data:labels.map(()=>objDiario||null),type:'line',borderColor:'#B5B0AD',borderDash:[5,3],pointRadius:0,fill:false}}
    ]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}}}}}}}},scales:{{y:{{ticks:{{callback:v=>'$'+(v/1000).toFixed(0)+'K',font:{{size:11}}}},grid:{{color:'#F2F1F0'}}}}}}}}
  }});
  if(charts.diaTicket)charts.diaTicket.destroy();
  charts.diaTicket=new Chart(document.getElementById('chartDiaTicket').getContext('2d'),{{
    type:'bar',
    data:{{labels,datasets:[{{label:'Ticket',data:tickets,backgroundColor:tickets.map(v=>v===maxT?'#1A7A4A':color+'88')}}]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{y:{{ticks:{{callback:v=>'$'+v.toFixed(0),font:{{size:11}}}},grid:{{color:'#F2F1F0'}}}}}}}}
  }});
  if(charts.diaClientes)charts.diaClientes.destroy();
  charts.diaClientes=new Chart(document.getElementById('chartDiaClientes').getContext('2d'),{{
    type:'bar',
    data:{{labels,datasets:[{{label:'Clientes',data:clientes,backgroundColor:clientes.map(v=>v===maxC?'#D4860A':color+'88')}}]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{y:{{ticks:{{font:{{size:11}}}},grid:{{color:'#F2F1F0'}}}}}}}}
  }});
}}
function fillTablaDias(u, mesesFiltro){{
  document.getElementById('tabla-dias-title').textContent=u+' · Análisis por Día de Semana';
  const ad=calcAnalisisdDia(u, mesesFiltro);
  const maxVenta=Math.max(...ad.map(d=>d.prom_venta));
  document.getElementById('tabla-dias-body').innerHTML=ad.map(d=>{{
    const esMejor=d.prom_venta===maxVenta;
    const c=esMejor?'color:#ED2E38;font-weight:700':'';
    const tend=tendenciaDia(d.serie);
    return `<tr>
      <td style="${{c}}">${{d.dia}}${{esMejor?' ⭐':''}}</td>
      <td style="${{c}}">${{fmt(d.prom_venta)}}</td>
      <td>${{fmtDec(d.prom_ticket)}}</td>
      <td>${{Math.round(d.prom_clientes)}}</td>
      <td style="color:#B5B0AD">${{d.n_dias}} días</td>
      <td>${{tendenciaBadgeHtml(tend)}}</td>
    </tr>`;
  }}).join('');
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
