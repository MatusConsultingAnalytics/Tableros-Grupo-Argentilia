"""
Pestaña PLAN DE RECUPERACIÓN del Tablero Ejecutivo — Grupo Gastronómico Argentilia.

Simulador de cheque promedio construido sobre el mismo modelo de la pestaña
Rentabilidad (rentabilidad.Modelo): estados de resultados de Contraloría.

Reglas de cálculo (validadas con Nat, 30/09/2026):
  · Solo se mueve el cheque promedio; los comensales quedan como en el mes base.
  · Venta proyectada    = comensales × (cheque real + ajuste).
  · Utilidad proyectada = utilidad real + venta adicional × (1 − % costo de venta de la unidad).
  · Meta = presupuesto de Contraloría del mismo mes.
  · SIN SUBSIDIO: el excedente de una unidad sobre su meta NO compensa el faltante de otra.
      cumplimiento grupo = Σ mín(valor, meta) ÷ Σ meta ; brecha grupo = Σ máx(0, meta − valor).
  · "Cheque para meta" = ajuste exacto por unidad = ⌈presupuesto de venta ÷ comensales reales − cheque real⌉, mínimo 0.
  · Presupuesto de utilidad inconsistente (menor a −20% de la venta presupuestada) → "meta de utilidad por validar"
    y se excluye de los totales de utilidad.

El cálculo corre en el navegador: el ajuste se escribe directamente en el tablero.
"""
import json

import rentabilidad as rt

AJUSTE_INICIAL = 15          # ajuste con el que abre la pestaña ($ por comensal)
UMBRAL_TICKET_ALCANZABLE = 0.08   # usado en el texto de la sección 05 (ver JS)


def _fecha(d):
    return d.strftime("%d/%m/%Y")


def datos_plan(base):
    """Estructura compacta por mes cerrado para el JavaScript."""
    mod = rt.Modelo(base)
    periodos = []
    for i in range(mod.mes_grupo):
        unidades = []
        for n in mod.orden:
            d = mod.u[n]
            venta, com = d["real"]["ventas"][i], d["real"]["comensales"][i]
            if venta <= 0 or com <= 0:
                continue
            pV = d["ppto"]["ventas"][i] if d["ppto"]["ventas"][i] > 0 else None
            pU = d["ppto"]["utilidad"][i] if pV else None
            pCom = d["ppto"]["comensales"][i] if pV and d["ppto"]["comensales"][i] > 0 else None
            nota = []
            if n in rt.APERTURAS:
                nota.append("Abrió " + _fecha(rt.APERTURAS[n]))
            if not pV:
                nota.append("sin presupuesto válido")
            elif venta / pV > rt.CUMPL_POR_VALIDAR:
                nota.append("presupuesto por validar")
            if pV and pU is not None and pU < -0.2 * pV:
                pU = None
                nota.append("meta de utilidad por validar")
            unidades.append({
                "id": rt._norm(n).lower().replace(" ", "-"), "nombre": n,
                "corto": rt.NOMBRE_CORTO.get(n, n), "nota": " · ".join(nota),
                "com": com, "venta": venta, "util": d["real"]["utilidad"][i],
                "costoPct": d["real"]["costo"][i] / venta,
                "pV": pV, "pU": pU, "pCom": pCom,
            })
        if unidades:
            periodos.append({"clave": f"m{i + 1}", "etiqueta": f"{rt.MESES_TIT[i]} {mod.anio}",
                             "predeterminado": False, "unidades": unidades})
    if periodos:
        periodos[-1]["predeterminado"] = True
    return {"anio": mod.anio, "periodos": periodos, "ajusteInicial": AJUSTE_INICIAL}


CSS = """
#plan-recuperacion{--pr-rojo:#ED2E38;--pr-rojo2:#C41E28;--pr-rojo3:#9E1820;--pr-osc:#3F3C40;--pr-txt:#2C2C2C;
 --pr-g1:#656266;--pr-g2:#77727A;--pr-g3:#B5B0AD;--pr-lin:#E7E4E1;--pr-fondo:#F4F2F0;--pr-verde:#2E6B4F;
 font-family:'IBM Plex Sans',Arial,sans-serif;color:var(--pr-txt);font-variant-numeric:tabular-nums;
 display:flex;flex-direction:column;gap:28px}
#plan-recuperacion *{box-sizing:border-box}
#plan-recuperacion{background:#F4F2F0;border-radius:12px;padding:28px 32px 36px}
#plan-recuperacion .pr-top{height:4px;background:#ED2E38;border-radius:2px;margin-bottom:-10px}
#plan-recuperacion .pr-kick{font-size:12px;letter-spacing:.14em;text-transform:uppercase;color:#77727A;font-weight:600}
#plan-recuperacion .pr-h1{margin:4px 0 2px;font-family:'IBM Plex Sans Condensed','IBM Plex Sans',sans-serif;font-size:40px;font-weight:700;color:#3F3C40}
#plan-recuperacion .pr-hs{font-size:15px;color:#656266}
@media print{#plan-recuperacion{background:#fff;padding:0}#plan-recuperacion .pr-ctrl select{display:none}}
#plan-recuperacion .pr-sec{display:flex;flex-direction:column;gap:14px}
#plan-recuperacion .pr-h{display:flex;align-items:baseline;gap:14px;border-bottom:1px solid #D6D2CF;padding-bottom:10px}
#plan-recuperacion .pr-h b{font-size:13px;color:var(--pr-rojo)}
#plan-recuperacion .pr-h h2{margin:0;font-size:20px;font-weight:600;color:var(--pr-osc)}
#plan-recuperacion .pr-h span{font-size:15px;color:var(--pr-g2)}
#plan-recuperacion .pr-card{background:#fff;border:1px solid var(--pr-lin);border-radius:12px;padding:22px 24px}
#plan-recuperacion .pr-palanca{display:flex;align-items:center;justify-content:space-between;gap:28px;flex-wrap:wrap}
#plan-recuperacion .pr-palanca p{margin:6px 0 0;font-size:14px;color:var(--pr-g1);line-height:1.5;max-width:560px}
#plan-recuperacion .pr-ctrl{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
#plan-recuperacion .pr-in{display:inline-flex;align-items:center;gap:4px;height:48px;border:2px solid var(--pr-osc);border-radius:8px;padding:0 12px;background:#fff}
#plan-recuperacion .pr-in input{width:72px;border:0;outline:0;font-family:inherit;font-size:20px;font-weight:700;background:transparent;color:var(--pr-txt)}
#plan-recuperacion .pr-btn{height:44px;min-width:64px;padding:0 14px;font-family:inherit;font-size:14px;font-weight:600;border-radius:8px;cursor:pointer;background:#fff;color:var(--pr-osc);border:1px solid #D6D2CF}
#plan-recuperacion .pr-btn.on{background:var(--pr-osc);color:#fff;border-color:var(--pr-osc)}
#plan-recuperacion .pr-btn.meta{color:var(--pr-rojo2);border:2px solid var(--pr-rojo2)}
#plan-recuperacion .pr-btn.meta.on{background:var(--pr-rojo2);color:#fff}
#plan-recuperacion select{height:44px;padding:0 12px;font-family:inherit;font-size:14px;border:1px solid #D6D2CF;border-radius:8px;background:#fff}
#plan-recuperacion .pr-grid4{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:16px}
#plan-recuperacion .pr-grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}
#plan-recuperacion .pr-kpi{display:flex;flex-direction:column;gap:12px}
#plan-recuperacion .pr-kt{font-size:14px;color:var(--pr-g1);font-weight:600;text-transform:uppercase;letter-spacing:.06em}
#plan-recuperacion .pr-kv{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap}
#plan-recuperacion .pr-kv .d{font-size:22px;color:#8A8589}
#plan-recuperacion .pr-kv .f{font-size:22px;color:var(--pr-rojo2);font-weight:700}
#plan-recuperacion .pr-kv .h{font-size:32px;font-weight:700;white-space:nowrap}
#plan-recuperacion .pr-pill{align-self:flex-start;font-size:14px;font-weight:700;padding:4px 10px;border-radius:999px}
#plan-recuperacion .pr-pill.up{background:#E4F0E9;color:var(--pr-verde)}
#plan-recuperacion .pr-pill.dn{background:#F8D7D9;color:var(--pr-rojo3)}
#plan-recuperacion .pr-bul{position:relative;height:44px}
#plan-recuperacion .pr-bul i{position:absolute;display:block}
#plan-recuperacion .pr-bul .tr{left:0;right:0;top:8px;height:28px;background:var(--pr-fondo);border-radius:4px}
#plan-recuperacion .pr-bul .r{left:0;top:10px;height:11px;background:var(--pr-g3);border-radius:2px}
#plan-recuperacion .pr-bul .p{left:0;top:23px;height:11px;background:var(--pr-rojo2);border-radius:2px}
#plan-recuperacion .pr-bul .m{top:2px;width:3px;height:40px;background:var(--pr-osc)}
#plan-recuperacion .pr-leg{display:flex;justify-content:space-between;font-size:12px;color:var(--pr-g2)}
#plan-recuperacion .pr-sw{display:inline-block;width:8px;height:8px;margin:0 4px 0 8px}
#plan-recuperacion .pr-cum{border-top:1px solid var(--pr-lin);padding-top:10px;font-size:14px;color:var(--pr-g1)}
#plan-recuperacion .pr-nota{font-size:12px;color:var(--pr-g2);line-height:1.4}
#plan-recuperacion .pr-franja{background:var(--pr-osc);color:#fff;border-radius:12px;padding:18px 24px;display:flex;align-items:center;gap:36px;flex-wrap:wrap}
#plan-recuperacion .pr-franja span{font-size:14px;color:#E7E4E1}
#plan-recuperacion .pr-franja .big{font-size:28px;font-weight:700;color:#fff}
#plan-recuperacion .pr-franja .was{font-size:20px;color:#CFCAC7}
#plan-recuperacion .pr-franja .to{font-size:20px;font-weight:700;color:#fff}
#plan-recuperacion table{width:100%;border-collapse:collapse;font-size:15px}
#plan-recuperacion th{background:var(--pr-rojo);color:#fff;font-weight:600;padding:14px 12px;text-align:right}
#plan-recuperacion th:first-child,#plan-recuperacion td:first-child{text-align:left;padding-left:18px}
#plan-recuperacion td{padding:12px;text-align:right;vertical-align:top}
#plan-recuperacion tbody tr:nth-child(odd){background:#FDF0F0}
#plan-recuperacion tbody tr:nth-child(even){background:#F5F5F5}
#plan-recuperacion tr.pr-tot{background:#fff!important;border-top:2px solid var(--pr-osc)}
#plan-recuperacion td small{display:block;font-size:12px;color:var(--pr-g2);font-weight:400}
#plan-recuperacion .pr-aj{display:inline-flex;align-items:center;height:44px;border:1px solid var(--pr-g3);border-radius:8px;background:#FFFBEA;padding:0 10px}
#plan-recuperacion .pr-aj input{width:60px;border:0;outline:0;background:transparent;font-family:inherit;font-size:16px;font-weight:700;text-align:right}
#plan-recuperacion .v{color:var(--pr-verde)} #plan-recuperacion .n{color:var(--pr-rojo2)} #plan-recuperacion .g{color:var(--pr-g1)}
#plan-recuperacion .pr-bar{display:flex;align-items:center;gap:12px;margin-top:12px}
#plan-recuperacion .pr-bar .l{width:120px;font-size:14px;font-weight:600;color:var(--pr-osc)}
#plan-recuperacion .pr-bar .t{position:relative;flex:1;height:34px;background:var(--pr-fondo);border-radius:4px}
#plan-recuperacion .pr-bar .t i{position:absolute;display:block;left:0;height:12px;border-radius:2px}
#plan-recuperacion .pr-bar .x{width:118px;text-align:right;font-size:13px}
#plan-recuperacion .pr-sr{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0)}
#plan-recuperacion .pr-pie{display:flex;gap:28px;font-size:13px;color:var(--pr-g1);line-height:1.5}
#plan-recuperacion .pr-pie b{flex-shrink:0;color:var(--pr-osc);text-transform:uppercase;letter-spacing:.08em;font-size:12px}
@media (max-width:1100px){#plan-recuperacion .pr-grid4{grid-template-columns:repeat(2,minmax(0,1fr))}
 #plan-recuperacion .pr-grid2{grid-template-columns:1fr}#plan-recuperacion .pr-tabla{overflow-x:auto}}
"""

HTML = """
<div class="pr-sec">
 <div class="pr-h"><b>01</b><h2>Palanca</h2><span>¿Cuánto sube el cheque promedio?</span></div>
 <div class="pr-card pr-palanca">
  <div><div style="font-size:17px;font-weight:600">Ajuste al cheque promedio por comensal</div>
   <p>Solo se modifica el cheque; los comensales quedan como en el mes base. Aplica un ajuste a todas las unidades,
   escribe uno por unidad en la tabla o usa <b>Cheque para meta</b>: el ajuste exacto que lleva a cada restaurante a su
   propia meta, sin que el excedente de una unidad compense a otra.</p></div>
  <div class="pr-ctrl">
   <label for="pr-periodo" style="font-size:14px;color:#656266">Mes base</label><select id="pr-periodo"></select>
   <label for="pr-global" style="font-size:14px;color:#656266;margin-left:8px">Todas las unidades</label>
   <span class="pr-in"><span style="font-size:18px;font-weight:600;color:#656266">+$</span><input id="pr-global" type="number" step="1"></span>
   <span id="pr-atajos" class="pr-ctrl"></span>
  </div>
 </div>
</div>
<div class="pr-sec">
 <div class="pr-h"><b>02</b><h2>Resultado proyectado del grupo</h2><span>De dónde estamos a dónde podemos llegar</span></div>
 <div id="pr-kpis" class="pr-grid4"></div>
 <div id="pr-franja" class="pr-franja"></div>
</div>
<div class="pr-sec">
 <div class="pr-h"><b>03</b><h2>Proyección por unidad</h2><span>Ajuste editable en cada renglón</span></div>
 <div class="pr-card pr-tabla" style="padding:0;overflow:hidden"><table>
  <thead><tr><th>Unidad</th><th>Comensales</th><th>Cheque resultado</th><th style="text-align:center">Ajuste ($)</th>
   <th>Cheque proyectado</th><th>Venta proyectada</th><th>Utilidad proyectada</th><th>Margen</th><th style="padding-right:18px">Brecha de cheque vs meta</th></tr></thead>
  <tbody id="pr-filas"></tbody></table></div>
</div>
<div class="pr-sec">
 <div class="pr-h"><b>04</b><h2>Cumplimiento de meta</h2><span>Resultado vs proyección</span></div>
 <div class="pr-grid2"><div class="pr-card" id="pr-barV"></div><div class="pr-card" id="pr-barU"></div></div>
</div>
<div class="pr-sec">
 <div class="pr-h"><b>05</b><h2>Conclusión</h2><span>El cheque promedio no es suficiente</span></div>
 <div class="pr-card" style="border-left:0;display:flex;flex-direction:column;gap:16px">
  <div style="display:flex;gap:18px;align-items:flex-start">
   <div style="width:6px;align-self:stretch;background:#C41E28;border-radius:3px;flex-shrink:0"></div>
   <div><div style="font-size:19px;font-weight:700;color:#3F3C40">Las acciones de incremento de ticket, por sí solas, no son suficientes.</div>
   <p style="margin:6px 0 0;font-size:15px;color:#656266;line-height:1.55">Cada unidad de negocio requiere una estrategia integral que combine
   afluencia (comensales), cheque promedio y mezcla de venta, costo de venta, productividad de nómina y control de gastos.
   Subir el cheque sin recuperar comensales ni corregir la estructura de costos no cierra la brecha contra la meta.</p></div>
  </div>
  <div id="pr-diag" class="pr-grid2" style="grid-template-columns:repeat(auto-fit,minmax(240px,1fr))"></div>
  <div class="pr-nota">Diagnóstico por unidad calculado con el mes base seleccionado (inferencia; por validar con cada gerencia).
   Cheque alcanzable = brecha de hasta 8% sobre el cheque real.</div>
 </div>
</div>
<div class="pr-card pr-pie"><b>Cómo se calcula</b><div>Venta proyectada = comensales del mes × (cheque resultado + ajuste).
 Utilidad proyectada = utilidad del mes + venta adicional × (1 − % costo de venta de la unidad); nómina, renta y gastos se mantienen fijos.
 Brecha de cheque = cheque real − cheque requerido para la meta de venta con los comensales reales. Sin compensar = el excedente de una unidad
 sobre su meta no cubre el faltante de otra (cumplimiento = Σ mín(valor, meta) ÷ Σ meta). Las barras del apartado 02 usan escala ajustada
 (inician en 85% del menor valor). Proyección = estimación; resultado y meta = dato confirmado (estados de resultados de Contraloría).</div></div>
"""

JS = r"""
(function () {
  var D = __DATOS__, raiz = document.getElementById('plan-recuperacion');
  if (!raiz || !D.periodos.length) { if (raiz) raiz.innerHTML = '<p>Sin estados financieros cerrados.</p>'; return; }
  var $ = function (id) { return document.getElementById(id); };
  var S = { per: D.periodos.findIndex(function (p) { return p.predeterminado; }), global: D.ajusteInicial, aj: {} };
  function num(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
  function miles(v) { return Math.round(v).toLocaleString('es-MX'); }
  function money(v) { return (v < 0 ? '−$' : '$') + Math.round(Math.abs(v)).toLocaleString('es-MX'); }
  function mill(v) { return (v < 0 ? '−$' : '$') + (Math.abs(v) / 1e6).toFixed(2) + 'M'; }
  function kk(v, sg) { var a = Math.abs(v), s = a >= 1e6 ? (a / 1e6).toFixed(2) + 'M' : (a / 1e3).toFixed(1) + 'K'; return (v < 0 ? '−$' : (sg ? '+$' : '$')) + s; }
  function pct(v) { return (v > 0 ? '+' : v < 0 ? '−' : '') + Math.abs(v * 100).toFixed(1) + '%'; }
  function p1(v) { return (v < 0 ? '−' : '') + Math.abs(v * 100).toFixed(1) + '%'; }
  function cl(v) { return v > 0.0005 ? 'v' : v < -0.0005 ? 'n' : 'g'; }
  function w15(r) { return (Math.max(0, Math.min(1.5, r)) / 1.5 * 100).toFixed(1) + '%'; }
  function esc(t) { return String(t).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function unidades() { return D.periodos[S.per].unidades; }
  function resetAj(v) { S.aj = {}; unidades().forEach(function (u) { S.aj[u.id] = v; }); }

  function calcular() {
    var G = { com: 0, venta: 0, util: 0, ventaP: 0, utilP: 0, pV: 0, pU: 0, pCom: 0, sV: 0, sVP: 0, sU: 0, sUP: 0, okV: 0, okVP: 0, okU: 0, okUP: 0, utilM: 0, utilPM: 0, sinMetaU: false, meta: {} };
    var filas = unidades().map(function (u) {
      var aj = num(S.aj[u.id]), tR = u.venta / u.com, tP = tR + aj;
      var vP = u.com * tP, dV = vP - u.venta, uP = u.util + dV * (1 - u.costoPct);
      var f = { u: u, aj: aj, tR: tR, tP: tP, vP: vP, dV: dV, uP: uP, dU: uP - u.util, conMeta: u.pV > 0 };
      G.com += u.com; G.venta += u.venta; G.util += u.util; G.ventaP += vP; G.utilP += uP;
      if (f.conMeta) {
        f.tM = u.pCom ? u.pV / u.pCom : null; f.tNec = u.pV / u.com; f.falt = f.tNec - tR;
        G.meta[u.id] = Math.max(0, Math.ceil(f.falt));
        G.pV += u.pV; G.pCom += u.pCom || 0;
        G.sV += Math.max(0, u.pV - u.venta); G.sVP += Math.max(0, u.pV - vP);
        if (u.pU != null) { G.pU += u.pU; G.sU += Math.max(0, u.pU - u.util); G.sUP += Math.max(0, u.pU - uP); G.okU += Math.min(u.util, u.pU); G.okUP += Math.min(uP, u.pU); G.utilM += u.util; G.utilPM += uP; } else { G.sinMetaU = true; }
        G.okV += Math.min(u.venta, u.pV); G.okVP += Math.min(vP, u.pV);
      }
      return f;
    });
    return { F: filas, G: G };
  }

  function pintarControles() {
    var sel = $('pr-periodo');
    sel.innerHTML = D.periodos.map(function (p, i) { return '<option value="' + i + '"' + (i === S.per ? ' selected' : '') + '>' + esc(p.etiqueta) + '</option>'; }).join('');
    var gi = $('pr-global'); if (document.activeElement !== gi) gi.value = S.global === 'meta' ? '' : S.global;
    var html = [0, 15, 30, 50].map(function (v) {
      return '<button type="button" class="pr-btn' + (String(S.global) === String(v) ? ' on' : '') + '" data-v="' + v + '">+$' + v + '</button>';
    }).join('') + '<button type="button" class="pr-btn meta' + (S.global === 'meta' ? ' on' : '') + '" data-v="meta">Cheque para meta</button>';
    $('pr-atajos').innerHTML = html;
  }

  function kpi(t, r, p, m, fmt, delta, up, mLbl, cLbl, cR, cP, cOk, nota) {
    var lo = Math.min(r, p, m) * 0.85, hi = Math.max(r, p, m) * 1.04;
    function wp(v) { return (Math.max(0, Math.min(1, (v - lo) / (hi - lo))) * 100).toFixed(1) + '%'; }
    return '<div class="pr-card pr-kpi"><div class="pr-kt">' + t + '</div>' +
      '<div class="pr-kv"><span class="d">' + fmt(r) + '</span><span class="f">→</span><span class="h">' + fmt(p) + '</span></div>' +
      '<div class="pr-pill ' + (up ? 'up' : 'dn') + '">' + delta + '</div>' +
      '<div><div class="pr-bul"><i class="tr"></i><i class="r" style="width:' + wp(r) + '"></i><i class="p" style="width:' + wp(p) + '"></i><i class="m" style="left:' + wp(m) + '"></i></div>' +
      '<div class="pr-leg"><span><i class="pr-sw" style="background:#B5B0AD;margin-left:0"></i>Resultado<i class="pr-sw" style="background:#C41E28"></i>Proyección</span>' +
      '<b style="color:#3F3C40">▍' + mLbl + ' ' + fmt(m) + '</b></div></div>' +
      '<div class="pr-cum">' + cLbl + ' <span style="color:#8A8589">' + cR + '</span> → <b class="' + (cOk ? 'v' : 'n') + '">' + cP + '</b></div>' +
      '<div class="pr-nota">' + nota + '</div></div>';
  }

  function pintar() {
    var R = calcular(), F = R.F, G = R.G;
    var tR = G.venta / G.com, tP = G.ventaP / G.com, falt = G.sV / G.com, faltP = G.sVP / G.com, tReq = tR + falt;
    var mR = G.util / G.venta, mP = G.utilP / G.ventaP, pVU = 0; F.forEach(function (f) { if (f.conMeta && f.u.pU != null) pVU += f.u.pV; }); var mM = pVU ? G.pU / pVU : 0;
    var pts = function (v) { return (v >= 0 ? '+' : '−') + Math.abs(v * 100).toFixed(1) + ' pts'; };
    $('pr-kpis').innerHTML =
      kpi('Cheque promedio', tR, tP, tReq, money, (tP >= tR ? '+' : '−') + money(Math.abs(tP - tR)) + ' por comensal', tP >= tR, 'Requerido',
        'Brecha sin compensar:', (falt > 0 ? '−' + money(falt) : money(0)), (faltP > 0.5 ? '−' + money(faltP) : money(0)), faltP < 0.5,
        'Requerido = cheque que lleva a cada unidad a su meta; el excedente de una no cubre a otra.') +
      kpi('Venta', G.venta, G.ventaP, G.pV, mill, kk(G.ventaP - G.venta, true) + ' · ' + pct(G.ventaP / G.venta - 1), G.ventaP >= G.venta, 'Meta',
        'Cumplimiento sin compensar:', p1(G.okV / G.pV), p1(G.okVP / G.pV), G.sVP < 0.5,
        'Con compensación entre unidades: ' + p1(G.venta / G.pV) + ' → ' + p1(G.ventaP / G.pV)) +
      kpi('Utilidad', G.util, G.utilP, G.pU, mill, kk(G.utilP - G.util, true) + ' · ' + pct(G.utilP / G.util - 1), G.utilP >= G.util, 'Meta',
        'Cumplimiento sin compensar:', p1(G.okU / G.pU), p1(G.okUP / G.pU), G.sUP < 0.5,
        'Con compensación entre unidades: ' + p1(G.utilM / G.pU) + ' → ' + p1(G.utilPM / G.pU) + (G.sinMetaU ? '. Excluye unidades con meta de utilidad por validar.' : '')) +
      kpi('Margen', mR, mP, mM, p1, pts(mP - mR), mP >= mR, 'Meta', 'vs meta:', pts(mR - mM), pts(mP - mM), mP >= mM,
        'Margen del grupo (utilidad total ÷ venta total).');
    var br = function (v) { return v > 0.5 ? '−' + kk(v).slice(1) : '$0'; };
    $('pr-franja').innerHTML =
      '<div><span>Utilidad adicional por mes </span><b class="big">' + kk(G.utilP - G.util, true) + '</b></div>' +
      '<div><span>Brecha de venta sin compensar </span><b class="was">' + br(G.sV) + '</b> <b style="color:#FF8A8F">→</b> <b class="to">' + br(G.sVP) + '</b></div>' +
      '<div><span>Brecha de utilidad sin compensar </span><b class="was">' + br(G.sU) + '</b> <b style="color:#FF8A8F">→</b> <b class="to">' + br(G.sUP) + '</b></div>';

    var filas = F.map(function (f) {
      var u = f.u, m = f.conMeta;
      var cV = m ? f.vP / u.pV - 1 : 0, cU = m && u.pU ? (f.uP - u.pU) / Math.abs(u.pU) : 0;
      return '<tr><td><b style="color:#3F3C40">' + esc(u.nombre) + '</b><small>' + esc(u.nota) + '</small></td>' +
        '<td>' + miles(u.com) + '</td>' +
        '<td>' + money(f.tR) + '<small>' + (f.tM ? 'meta ' + money(f.tM) : '—') + '</small></td>' +
        '<td style="text-align:center"><label class="pr-sr" for="pr-aj-' + u.id + '">Ajuste ' + esc(u.nombre) + '</label>' +
        '<span class="pr-aj">$<input id="pr-aj-' + u.id + '" data-id="' + u.id + '" type="number" step="1" value="' + esc(S.aj[u.id]) + '"></span></td>' +
        '<td><b>' + money(f.tP) + '</b>' + (f.tM ? '<small class="' + (f.tP >= f.tM ? 'v' : 'n') + '">' + (f.tP >= f.tM ? 'meta alcanzada (+' + money(f.tP - f.tM) + ')' : 'faltan ' + money(f.tM - f.tP) + ' para meta') + '</small>' : '') + '</td>' +
        '<td><b>' + mill(f.vP) + '</b><small>resultado ' + mill(u.venta) + ' · <span class="' + cl(f.dV) + '">' + kk(f.dV, true) + '</span></small>' +
        (m ? '<small>meta ' + mill(u.pV) + ' · <b class="' + cl(cV) + '">' + pct(cV) + '</b></small>' : '<small>sin presupuesto</small>') + '</td>' +
        '<td><b class="' + (f.uP < 0 ? 'n' : '') + '">' + mill(f.uP) + '</b><small>resultado ' + mill(u.util) + ' · <span class="' + cl(f.dU) + '">' + kk(f.dU, true) + '</span></small>' +
        (m && u.pU != null ? '<small>meta ' + mill(u.pU) + ' · <b class="' + cl(cU) + '">' + pct(cU) + '</b></small>' : '') + '</td>' +
        '<td><b>' + (f.uP / f.vP * 100).toFixed(1) + '%</b><small>resultado ' + (u.util / u.venta * 100).toFixed(1) + '%</small></td>' +
        '<td style="padding-right:18px">' + (m ? '<b style="font-size:17px" class="' + (f.falt > 0 ? 'n' : 'v') + '">' + (f.falt > 0 ? '−' + money(f.falt) : '+' + money(-f.falt)) + '</b>' +
          '<small>requerido ' + money(f.tNec) + ' vs real ' + money(f.tR) + '</small><small>' +
          (f.falt > 0 ? (f.aj > 0 ? 'el ajuste cubre ' + Math.min(100, Math.round(f.aj / f.falt * 100)) + '%' : 'sin ajuste') : 'ya supera la meta') + '</small>' : '—') + '</td></tr>';
    }).join('');
    filas += '<tr class="pr-tot"><td><b>Grupo</b></td><td><b>' + miles(G.com) + '</b></td><td><b>' + money(tR) + '</b></td>' +
      '<td style="text-align:center"><b>' + (tP >= tR ? '+' : '−') + money(Math.abs(tP - tR)) + ' prom.</b></td><td><b>' + money(tP) + '</b></td>' +
      '<td><b>' + mill(G.ventaP) + '</b><small>meta ' + mill(G.pV) + ' · <b class="' + (G.sVP < 0.5 ? 'v' : 'n') + '">' + p1(G.okVP / G.pV) + '</b> sin compensar</small></td>' +
      '<td><b>' + mill(G.utilP) + '</b><small>meta ' + mill(G.pU) + ' · <b class="' + (G.sUP < 0.5 ? 'v' : 'n') + '">' + p1(G.okUP / G.pU) + '</b> sin compensar</small></td>' +
      '<td><b>' + p1(mP) + '</b></td>' +
      '<td style="padding-right:18px"><b style="font-size:17px" class="' + (falt > 0 ? 'n' : 'v') + '">' + (falt > 0 ? '−' + money(falt) : money(0)) + '</b>' +
      '<small>requerido ' + money(tReq) + ' vs real ' + money(tR) + '</small><small>' +
      (G.sV > 0 ? 'el ajuste cubre ' + Math.round((G.sV - G.sVP) / G.sV * 100) + '% · sin compensar' : 'todas en meta') + '</small></td></tr>';
    var tb = $('pr-filas'), act = document.activeElement;
    if (act && act.dataset && act.dataset.id && tb.contains(act) && tb.rows.length) {
      // Se está escribiendo en un ajuste: se actualizan las demás celdas sin tocar el campo activo.
      var tmp = document.createElement('tbody'); tmp.innerHTML = filas;
      for (var i = 0; i < tmp.rows.length && i < tb.rows.length; i++) {
        for (var j = 0; j < tmp.rows[i].cells.length; j++) {
          var celda = tb.rows[i].cells[j];
          if (celda && !celda.querySelector('input')) celda.innerHTML = tmp.rows[i].cells[j].innerHTML;
        }
      }
    } else { tb.innerHTML = filas; }

    function barras(titulo, clave, nota) {
      return '<div style="display:flex;justify-content:space-between;align-items:center"><div style="font-size:14px;color:#656266;font-weight:500">' + titulo + '</div>' +
        '<div class="pr-leg" style="gap:8px"><span><i class="pr-sw" style="background:#B5B0AD"></i>Resultado<i class="pr-sw" style="background:#C41E28"></i>Proyección</span></div></div>' +
        F.filter(function (f) { return f.conMeta && (clave === 'V' || f.u.pU != null); }).map(function (f) {
          var r = clave === 'V' ? f.u.venta / f.u.pV : f.u.util / f.u.pU, p = clave === 'V' ? f.vP / f.u.pV : f.uP / f.u.pU;
          return '<div class="pr-bar"><div class="l">' + esc(f.u.corto) + '</div><div class="t"><i style="top:4px;background:#B5B0AD;width:' + w15(r) + '"></i>' +
            '<i style="top:18px;background:#C41E28;width:' + w15(p) + '"></i><i style="left:66.67%;top:0;width:2px;height:34px;background:#3F3C40;border-radius:0"></i></div>' +
            '<div class="x"><span style="color:#77727A">' + Math.round(r * 100) + '%</span> → <b>' + Math.round(p * 100) + '%</b></div></div>';
        }).join('') + '<div class="pr-nota" style="margin-top:12px">' + nota + '</div>';
    }
    $('pr-diag').innerHTML = F.filter(function (f) { return f.conMeta; }).map(function (f) {
      var u = f.u, gap = f.falt / f.tR, frentes = [];
      var uEnMeta = u.util + Math.max(0, u.pV - u.venta) * (1 - u.costoPct);
      if (f.falt <= 0) frentes.push('Supera su meta de venta: sostener y documentar prácticas' + (u.nota.indexOf('por validar') >= 0 ? ' (presupuesto por validar)' : ''));
      else if (gap <= 0.08) frentes.push('Brecha de cheque de ' + money(f.falt) + ' (+' + Math.round(gap * 100) + '%): alcanzable con venta sugerida y mezcla de bebidas');
      else if (u.pCom && u.com < u.pCom * 0.95) frentes.push('Brecha de cheque de ' + money(f.falt) + ' (+' + Math.round(gap * 100) + '%): no es realista solo con ticket; recuperar afluencia');
      else frentes.push('Brecha de cheque de ' + money(f.falt) + ' (+' + Math.round(gap * 100) + '%) con afluencia en línea: revisar precios, menú y mezcla de venta');
      if (u.pCom && u.com < u.pCom) frentes.push('Comensales ' + Math.round((1 - u.com / u.pCom) * 100) + '% bajo lo presupuestado (' + miles(u.com) + ' vs ' + miles(u.pCom) + ')');
      if (u.pU != null && f.falt > 0 && uEnMeta < u.pU) frentes.push('Aun cumpliendo la venta, la utilidad quedaría ' + kk(uEnMeta - u.pU) + ' bajo meta: revisar nómina y gastos');
      var nivel = f.falt <= 0 ? ['#2E6B4F', 'Sostener'] : (gap <= 0.08 && !(u.pU != null && uEnMeta < u.pU)) ? ['#8A6A1C', 'Ticket + mezcla'] : ['#C41E28', 'Estrategia integral'];
      return '<div style="border:1px solid #E7E4E1;border-radius:10px;padding:14px 16px;display:flex;flex-direction:column;gap:8px">' +
        '<div style="display:flex;justify-content:space-between;align-items:center;gap:8px"><b style="color:#3F3C40">' + esc(u.corto) + '</b>' +
        '<span style="font-size:12px;font-weight:700;color:' + nivel[0] + ';border:1px solid ' + nivel[0] + ';border-radius:999px;padding:2px 8px;white-space:nowrap">' + nivel[1] + '</span></div>' +
        frentes.map(function (t) { return '<div style="font-size:13px;color:#656266;line-height:1.45">· ' + t + '</div>'; }).join('') + '</div>';
    }).join('');
    $('pr-barV').innerHTML = barras('Venta como % de la meta', 'V', 'Línea vertical = 100% de la meta. Escala hasta 150%.');
    $('pr-barU').innerHTML = barras('Utilidad como % de la meta', 'U', 'Línea vertical = 100% de la meta. Utilidad negativa se muestra en 0%.');
    return G;
  }

  raiz.addEventListener('change', function (e) {
    if (e.target.id === 'pr-periodo') {
      var previo = S.aj; S.per = +e.target.value;
      if (S.global === 'meta') { resetAj(0); S.aj = Object.assign({}, calcular().G.meta); }
      else { var base = S.global === '' ? D.ajusteInicial : S.global; resetAj(base); unidades().forEach(function (u) { if (previo[u.id] != null) S.aj[u.id] = previo[u.id]; }); }
      pintarControles(); pintar();
    }
  });
  raiz.addEventListener('input', function (e) {
    if (e.target.id === 'pr-global') { S.global = e.target.value; resetAj(e.target.value); pintarControles(); pintar(); }
    else if (e.target.dataset && e.target.dataset.id) {
      S.aj[e.target.dataset.id] = e.target.value; S.global = ''; pintarControles(); pintar();
    }
  });
  raiz.addEventListener('click', function (e) {
    var b = e.target.closest ? e.target.closest('.pr-btn') : null; if (!b) return;
    var v = b.getAttribute('data-v');
    if (v === 'meta') { S.global = 'meta'; resetAj(0); S.aj = Object.assign({}, calcular().G.meta); }
    else { S.global = +v; resetAj(+v); }
    pintarControles(); pintar();
  });
  resetAj(S.global); pintarControles(); pintar();
})();
"""



def generar_seccion(base):
    """HTML completo de la pestaña (div#plan.section)."""
    cab = ('<div class="pr-top"></div><div><div class="pr-kick">Simulador · estados financieros de Contraloría</div>'
           '<h1 class="pr-h1">Plan de recuperación</h1>'
           '<div class="pr-hs">Proyección modificando solo el cheque promedio · comparación contra resultado y meta</div></div>')
    if not base or base.get("error"):
        msg = rt.esc(base.get("error") if base else "Sin información financiera.")
        return (f'<div id="plan" class="section"><style>{CSS}</style><div id="plan-recuperacion">{cab}'
                f'<div class="pr-card" style="color:#9E1820">La información financiera no está disponible en esta actualización. '
                f'{msg}</div></div></div>')
    d = datos_plan(base)
    js = JS.replace("__DATOS__", json.dumps(d, ensure_ascii=False).replace("</", "<\\/"))
    return (f'<div id="plan" class="section"><style>{CSS}</style>'
            f'<div id="plan-recuperacion">{cab}{HTML}</div><script>{js}</script></div>')
