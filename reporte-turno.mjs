// Planta Aerobol · Resumen automático por correo al cierre de cada turno (o de la jornada).
// Lo ejecuta GitHub Actions (.github/workflows/reporte-turno.yml). Lee Firestore con una
// cuenta de servicio y envía el correo con Microsoft 365 (Microsoft Graph).
//
// Variables de entorno:
//   FIREBASE_SA          JSON de la cuenta de servicio de Firebase (secreto)
//   M365_TENANT_ID       Id. de directorio (inquilino) de Microsoft 365 (secreto)
//   M365_CLIENT_ID       Id. de aplicación (cliente) del registro de la app (secreto)
//   M365_CLIENT_SECRET   Valor del secreto de cliente de la app (secreto)
//   REPORTE_DE           Buzón que envía, p. ej. reportes@aerobol.com.co
//   REPORTE_PARA         Destinatarios separados por coma
//   REPORTE_CC           (opcional) Con copia, separados por coma
//   REPORTE_MODO         "turno" (un correo por turno) o "jornada" (un correo diario con día + noche)
//   REPORTE_PRUEBA       (opcional) Correo para envíos de prueba
//   FECHA, TURNO         (opcional) Forzar un turno: AAAA-MM-DD y D o N
//   SOLO_PRUEBA          "true" para enviar solo a REPORTE_PRUEBA
//   REPORTE_ACTIVO       "si" para que salgan los envíos programados; cualquier otro valor los apaga
//                        (las pruebas manuales desde Actions → Run workflow siempre funcionan)
//   REPORTE_SI_CERO      "si" para enviar aunque el turno no tenga producción (por defecto no se envía)
//   DRY_RUN              "1" para no enviar: guarda el correo en reporte.html
//   DATOS_PRUEBA         (solo pruebas locales) ruta a un JSON con {config, dias}

import fs from 'node:fs';

/* ---------- reglas del negocio (iguales a las del tablero) ---------- */
const UXC = 24;
const TURNOS = {
  D: { n: 'Día',   ini: 360,  fin: 1080, desc: [720, 780] },   // 6:00 a.m. – 6:00 p.m., almuerzo 12–1
  N: { n: 'Noche', ini: 1080, fin: 1800, desc: null },         // 6:00 p.m. – 6:00 a.m.
};
const ov = (a, b, c, d) => Math.max(0, Math.min(b, d) - Math.max(a, c));
const prodMinT = (a, b, t) => { const T = TURNOS[t]; return ov(a, b, T.ini, T.fin) - (T.desc ? ov(a, b, T.desc[0], T.desc[1]) : 0); };
const jor = t => prodMinT(0, 2880, t);
const medianoche = k => { const [y, m, d] = k.split('-').map(Number); return Date.UTC(y, m - 1, d, 5); };   // 00:00 Bogotá
const tm = (ts, k) => (ts - medianoche(k)) / 60000;
const bogMin = ts => { const d = new Date(ts - 5 * 3600000); return d.getUTCHours() * 60 + d.getUTCMinutes(); };
const bogKey = ts => new Date(ts - 5 * 3600000).toISOString().slice(0, 10);
const diaMas = (k, n) => { const [y, m, d] = k.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10); };
const evDe = (d, t) => (d?.eventos || []).filter(e => (e.turno || 'D') === t);

function normEvents(list) {
  const ev = (list || []).slice().sort((a, b) => a.ts - b.ts).map(e => ({ ...e }));
  const run = {};
  for (const e of ev) {
    if (e.tipo !== 'prod') continue;
    const prev = run[e.maq] || 0;
    if (e.acum != null) { e.unidades = Math.max(0, e.acum - prev); run[e.maq] = Math.max(prev, e.acum); }
    else run[e.maq] = prev + (e.unidades || 0);
    e.totalMaq = run[e.maq];
  }
  return ev;
}

/* ---------- formato ---------- */
const nf0 = new Intl.NumberFormat('es-CO', { maximumFractionDigits: 0 });
const nf1 = new Intl.NumberFormat('es-CO', { maximumFractionDigits: 1, minimumFractionDigits: 1 });
const hhmm = min => { min = ((min % 1440) + 1440) % 1440; let h = Math.floor(min / 60), m = Math.floor(min % 60); const ap = h < 12 ? 'a.m.' : 'p.m.'; h = h % 12 || 12; return `${h}:${String(m).padStart(2, '0')} ${ap}`; };
const horaTs = ts => hhmm(bogMin(ts));
const dur = min => min < 60 ? `${nf0.format(min)} min` : `${Math.floor(min / 60)} h ${Math.round(min % 60)} min`;
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fechaLarga = k => { const [y, m, d] = k.split('-').map(Number); const s = new Date(Date.UTC(y, m - 1, d)).toLocaleDateString('es-CO', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' }); return s.charAt(0).toUpperCase() + s.slice(1); };
const fechaCorta = k => { const [y, m, d] = k.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)).toLocaleDateString('es-CO', { weekday: 'short', day: 'numeric', month: 'short', timeZone: 'UTC' }); };

/* ---------- resumen de un turno ---------- */
function resumen(doc, fecha, t, cfg) {
  const T = TURNOS[t];
  const ev = normEvents(evDe(doc, t));
  const pers = doc?.personal?.[t] || {};
  const NOMBRE_MAQ = { m1: 'Máquina automática', m2: 'Máquina manual' };   // iguales a los del tablero
  const maqs = ((cfg.maquinas && cfg.maquinas.length) ? cfg.maquinas : [{ id: 'm1' }]).map(m => ({ ...m, nombre: NOMBRE_MAQ[m.id] || m.nombre }));
  const M = {};
  for (const m of maqs) M[m.id] = { id: m.id, nombre: m.nombre, und: 0, reportes: 0, fallas: [], downMin: 0, abierta: null, pers: pers[m.id] ?? null };
  const fallas = [];
  const quien = {};
  for (const e of ev) {
    const m = M[e.maq] || (M[e.maq] = { id: e.maq, nombre: e.maq, und: 0, reportes: 0, fallas: [], downMin: 0, abierta: null, pers: pers[e.maq] ?? null });
    if (e.por) quien[e.por] = (quien[e.por] || 0) + 1;
    if (e.tipo === 'prod') { m.und += e.unidades; m.reportes++; }
    else if (e.tipo === 'falla') { if (!m.abierta) { m.abierta = { ts: e.ts, nota: e.nota || '' }; } }
    else if (e.tipo === 'reanuda' && m.abierta) {
      const min = prodMinT(tm(m.abierta.ts, fecha), tm(e.ts, fecha), t);
      m.downMin += min; const f = { maq: m.nombre, ini: m.abierta.ts, fin: e.ts, min, nota: m.abierta.nota }; m.fallas.push(f); fallas.push(f); m.abierta = null;
    }
  }
  for (const m of Object.values(M)) if (m.abierta) {      // falla sin reanudar: cuenta hasta el fin del turno
    const min = prodMinT(tm(m.abierta.ts, fecha), T.fin, t);
    m.downMin += min; const f = { maq: m.nombre, ini: m.abierta.ts, fin: null, min, nota: m.abierta.nota }; m.fallas.push(f); fallas.push(f); m.abierta = null;
  }
  const J = jor(t);
  const lista = Object.values(M).map(m => {
    const eff = Math.max(0, J - m.downMin);
    return { ...m, cajas: Math.floor(m.und / UXC), sueltas: m.und % UXC, eff,
      envH: (m.und > 0 && eff >= 15) ? m.und / eff * 60 : null,
      envOp: (m.pers > 0 && m.und > 0) ? m.und / m.pers : null };
  });
  const total = lista.reduce((s, m) => s + m.und, 0);
  const persTot = lista.reduce((s, m) => s + (m.pers || 0), 0);
  const conPers = lista.some(m => m.pers != null);
  const meta = cfg.meta || null;
  return {
    fecha, t, nombre: T.n, horario: `${hhmm(T.ini)} – ${hhmm(T.fin)}`, J, total, cajas: Math.floor(total / UXC), sueltas: total % UXC,
    maquinas: lista, fallas: fallas.sort((a, b) => a.ini - b.ini), quien,
    reportes: ev.filter(e => e.tipo === 'prod').length,
    envH: total > 0 ? total / J * 60 : null,
    persTot: conPers ? persTot : null, envOp: conPers && persTot > 0 && total > 0 ? total / persTot : null,
    meta, pct: meta ? total / meta * 100 : null, dif: meta ? total - meta : null,
    downMin: lista.reduce((s, m) => s + m.downMin, 0),
  };
}

/* ---------- correo (HTML compatible con Outlook) ---------- */
const C = { navy: '#14284B', oro: '#9C7740', gris: '#66748A', linea: '#E2E7EE', fondo: '#F3F5F8', ok: '#2E8B57', alto: '#C0392B' };
const celda = (l, v, s = '') => `<td style="padding:14px 16px;border:1px solid ${C.linea};background:#fff;vertical-align:top;width:33%">
  <div style="font-size:12px;color:${C.gris};margin-bottom:4px">${l}</div>
  <div style="font-size:22px;font-weight:bold;color:${C.navy};line-height:1.2">${v}</div>
  ${s ? `<div style="font-size:12px;color:${C.gris};margin-top:4px">${s}</div>` : ''}</td>`;

function bloqueTurno(r, ant, conTitulo) {
  const metaTxt = r.meta
    ? `${nf0.format(Math.floor(r.pct))} %`
    : '—';
  const metaSub = r.meta
    ? (r.dif >= 0 ? `<span style="color:${C.ok};font-weight:bold">Meta cumplida</span> · +${nf0.format(r.dif)} sobre ${nf0.format(r.meta)}` : `Faltaron ${nf0.format(-r.dif)} de ${nf0.format(r.meta)}`)
    : 'Sin meta definida';
  const antTxt = ant && ant.total > 0
    ? `${nf0.format(ant.total)}`
    : '—';
  const antSub = ant && ant.total > 0
    ? `Turno ${ant.nombre.toLowerCase()} · ${fechaCorta(ant.fecha)} · ${r.total - ant.total >= 0 ? '+' : ''}${nf0.format(r.total - ant.total)} (${r.total - ant.total >= 0 ? '+' : ''}${nf1.format((r.total - ant.total) / ant.total * 100)} %)`
    : 'Sin registros';

  const filasMaq = r.maquinas.map(m => `<tr>
      <td style="padding:10px 12px;border-bottom:1px solid ${C.linea};font-weight:bold;color:${C.navy}">${esc(m.nombre)}</td>
      <td style="padding:10px 12px;border-bottom:1px solid ${C.linea};text-align:right">${nf0.format(m.und)}</td>
      <td style="padding:10px 12px;border-bottom:1px solid ${C.linea};text-align:right">${nf0.format(m.cajas)}${m.sueltas ? ` <span style="color:${C.gris}">+${m.sueltas}</span>` : ''}</td>
      <td style="padding:10px 12px;border-bottom:1px solid ${C.linea};text-align:right">${m.envH != null ? nf0.format(m.envH) : '—'}</td>
      <td style="padding:10px 12px;border-bottom:1px solid ${C.linea};text-align:right">${m.pers != null ? nf0.format(m.pers) : '—'}</td>
      <td style="padding:10px 12px;border-bottom:1px solid ${C.linea};text-align:right">${m.envOp != null ? nf0.format(m.envOp) : '—'}</td>
      <td style="padding:10px 12px;border-bottom:1px solid ${C.linea};text-align:right;white-space:nowrap;${m.downMin ? `color:${C.alto}` : ''}">${m.downMin ? dur(m.downMin) : '—'}</td>
    </tr>`).join('');
  const th = (t, a = 'right') => `<th style="padding:8px 12px;text-align:${a};font-size:12px;font-weight:normal;color:${C.gris};border-bottom:2px solid ${C.linea}">${t}</th>`;

  const fallasHtml = r.fallas.length
    ? `<div style="font-size:15px;font-weight:bold;color:${C.navy};margin:22px 0 8px">Paradas de máquina</div>
       <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px">
       ${r.fallas.map(f => `<tr><td style="padding:8px 12px;border-bottom:1px solid ${C.linea};color:${C.alto};font-weight:bold;white-space:nowrap">${horaTs(f.ini)} – ${f.fin ? horaTs(f.fin) : 'fin del turno'}</td>
         <td style="padding:8px 12px;border-bottom:1px solid ${C.linea}">${esc(f.maq)}${f.nota ? ` · ${esc(f.nota)}` : ''}</td>
         <td style="padding:8px 12px;border-bottom:1px solid ${C.linea};text-align:right;white-space:nowrap">${dur(f.min)} productivos</td></tr>`).join('')}
       </table>`
    : `<div style="font-size:13px;color:${C.ok};margin:18px 0 0">Sin paradas de máquina reportadas en el turno.</div>`;

  const quien = Object.entries(r.quien).map(([n, c]) => `${esc(n)} (${c})`).join(', ');

  return `
  ${conTitulo ? `<div style="font-size:18px;font-weight:bold;color:${C.navy};margin:30px 0 10px;padding-top:18px;border-top:2px solid ${C.linea}">Turno ${r.nombre.toLowerCase()} · ${r.horario}</div>` : ''}
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
    <tr>
      <td colspan="3" style="padding:20px 20px;border:1px solid ${C.linea};background:#fff">
        <div style="font-size:13px;color:${C.gris}">Producido en el turno</div>
        <div style="font-size:40px;font-weight:bold;color:${C.navy};line-height:1.15">${nf0.format(r.total)} <span style="font-size:16px;font-weight:normal;color:${C.gris}">envases</span></div>
        <div style="font-size:14px;color:${C.gris}"><b style="color:${C.navy}">${nf0.format(r.cajas)}</b> cajas de ${UXC}${r.sueltas ? ` + ${r.sueltas} envases` : ''}</div>
      </td>
    </tr>
    <tr>${celda('Cumplimiento de la meta', metaTxt, metaSub)}${celda('Envases por hora', r.envH != null ? nf0.format(r.envH) : '—', `Promedio en ${nf1.format(r.J / 60)} h productivas`)}${celda('Turno anterior', antTxt, antSub)}</tr>
    <tr>${celda('Operarios en el turno', r.persTot != null ? nf0.format(r.persTot) : '—', r.persTot == null ? 'No se registraron' : '')}${celda('Envases por operario', r.envOp != null ? nf0.format(r.envOp) : '—')}${celda('Tiempo de máquinas detenidas', r.downMin ? dur(r.downMin) : 'Sin paradas', r.fallas.length ? `${r.fallas.length} parada${r.fallas.length === 1 ? '' : 's'}` : '')}</tr>
  </table>
  <div style="font-size:15px;font-weight:bold;color:${C.navy};margin:22px 0 8px">Por máquina</div>
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:13px;color:#2B3A55">
    <tr>${th('Máquina', 'left')}${th('Envases')}${th('Cajas')}${th('Env/h')}${th('Operarios')}${th('Env/operario')}${th('Detenida')}</tr>
    ${filasMaq}
  </table>
  ${fallasHtml}
  <div style="font-size:12px;color:${C.gris};margin-top:16px">${r.reportes} reporte${r.reportes === 1 ? '' : 's'} de producción${quien ? ` · Movimientos registrados por: ${quien}` : ''}</div>`;
}

function correo(titulo, subtitulo, cuerpo) {
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>${esc(titulo)}</title></head>
<body style="margin:0;padding:0;background:${C.fondo};font-family:'Segoe UI',Arial,Helvetica,sans-serif;color:${C.navy}">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:${C.fondo}"><tr><td align="center" style="padding:24px 12px">
<table role="presentation" width="680" cellpadding="0" cellspacing="0" style="max-width:680px;width:100%;border-collapse:collapse">
  <tr><td style="background:${C.navy};padding:22px 28px">
    <div style="font-size:22px;letter-spacing:6px;color:#fff">AERO<span style="color:#C9A86F">BOL</span></div>
    <div style="font-size:12px;color:#B9C6D8;margin-top:2px">Planta Barranquilla · Espuma Torito</div>
  </td></tr>
  <tr><td style="background:#fff;padding:24px 28px 8px;border-left:1px solid ${C.linea};border-right:1px solid ${C.linea}">
    <div style="font-size:20px;font-weight:bold;color:${C.navy}">${esc(titulo)}</div>
    <div style="font-size:14px;color:${C.gris};margin-top:4px">${esc(subtitulo)}</div>
  </td></tr>
  <tr><td style="background:#fff;padding:12px 28px 28px;border:1px solid ${C.linea};border-top:0">${cuerpo}</td></tr>
  <tr><td style="padding:16px 4px;font-size:11px;color:${C.gris};text-align:center">Correo automático del Tablero de producción Aerobol · jeancar07.github.io/planta-aerobol<br>Las cifras corresponden a lo registrado en el tablero por el personal de planta.</td></tr>
</table></td></tr></table></body></html>`;
}

/* ---------- datos ---------- */
async function abrirDatos() {
  if (process.env.DATOS_PRUEBA) {
    const j = JSON.parse(fs.readFileSync(process.env.DATOS_PRUEBA, 'utf8'));
    return { config: async () => j.config, dia: async k => j.dias[k] || null };
  }
  const admin = (await import('firebase-admin')).default;
  admin.initializeApp({ credential: admin.credential.cert(JSON.parse(process.env.FIREBASE_SA)) });
  const db = admin.firestore();
  return {
    config: async () => (await db.doc('config/planta').get()).data() || {},
    dia: async k => { const s = await db.doc('dias/' + k).get(); return s.exists ? s.data() : null; },
  };
}

/* ---------- envío por Microsoft Graph ---------- */
async function enviar({ asunto, html, para, cc }) {
  const tok = await fetch(`https://login.microsoftonline.com/${process.env.M365_TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: process.env.M365_CLIENT_ID, client_secret: process.env.M365_CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
  }).then(r => r.json());
  if (!tok.access_token) throw new Error('Microsoft 365 no entregó el token: ' + (tok.error_description || JSON.stringify(tok)));
  const lista = s => s.split(/[,;\s]+/).filter(Boolean).map(a => ({ emailAddress: { address: a } }));
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(process.env.REPORTE_DE)}/sendMail`, {
    method: 'POST', headers: { Authorization: 'Bearer ' + tok.access_token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ message: { subject: asunto, body: { contentType: 'HTML', content: html }, toRecipients: lista(para), ccRecipients: cc ? lista(cc) : [] }, saveToSentItems: true }),
  });
  if (r.status !== 202) throw new Error(`Microsoft 365 rechazó el envío (${r.status}): ${await r.text()}`);
}

/* ---------- principal ---------- */
async function main() {
  const programado = process.env.GITHUB_EVENT_NAME === 'schedule';
  if (programado && String(process.env.REPORTE_ACTIVO || '').trim().toLowerCase() !== 'si') {
    console.log('Envíos programados APAGADOS (variable REPORTE_ACTIVO distinta de "si"). No se envía nada.');
    return;
  }
  const modo = (process.env.REPORTE_MODO || 'turno').toLowerCase();
  const manual = !!(process.env.FECHA && process.env.TURNO);
  const ahora = process.env.AHORA ? Date.parse(process.env.AHORA) : Date.now();   // AHORA solo para pruebas
  // turno que acaba de cerrar
  let fecha, t;
  if (manual) { fecha = process.env.FECHA; t = process.env.TURNO.toUpperCase().startsWith('N') ? 'N' : 'D'; }
  else {
    const m = bogMin(ahora), k = bogKey(ahora);
    if (m >= 1080) { fecha = k; t = 'D'; }               // después de 6 p.m.: cerró el turno día de hoy
    else if (m < 360) { fecha = diaMas(k, -1); t = 'D'; } // madrugada: el último cerrado es el día de ayer
    else { fecha = diaMas(k, -1); t = 'N'; }              // después de 6 a.m.: cerró la noche que empezó ayer
  }
  if (modo === 'jornada' && !manual && t === 'D') { console.log('Modo jornada: el resumen se envía al cerrar el turno noche. Nada que enviar ahora.'); return; }

  const datos = await abrirDatos();
  const cfg = await datos.config();
  const doc = await datos.dia(fecha);

  let asunto, html, totalEnviado = 0;
  if (modo === 'jornada') {
    const rD = resumen(doc, fecha, 'D', cfg), rN = resumen(doc, fecha, 'N', cfg);
    const docAnt = await datos.dia(diaMas(fecha, -1));
    const antN = resumen(docAnt, diaMas(fecha, -1), 'N', cfg);
    const tot = rD.total + rN.total; totalEnviado = tot; const metaJ = cfg.meta ? cfg.meta * 2 : null;
    asunto = `Planta Aerobol · Resumen de la jornada · ${fechaCorta(fecha)} · ${nf0.format(tot)} envases${metaJ ? ` (${nf0.format(Math.floor(tot / metaJ * 100))} % de la meta)` : ''}`;
    const cab = `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:8px"><tr>
      ${celda('Total de la jornada', `${nf0.format(tot)} envases`, `${nf0.format(Math.floor(tot / UXC))} cajas de ${UXC}`)}
      ${celda('Turno día', nf0.format(rD.total), rD.meta ? `${nf0.format(Math.floor(rD.pct))} % de la meta` : '')}
      ${celda('Turno noche', nf0.format(rN.total), rN.meta ? `${nf0.format(Math.floor(rN.pct))} % de la meta` : '')}</tr></table>`;
    html = correo('Resumen de la jornada', `${fechaLarga(fecha)} · turno día y turno noche`, cab + bloqueTurno(rD, antN, true) + bloqueTurno(rN, rD, true));
  } else {
    const r = resumen(doc, fecha, t, cfg); totalEnviado = r.total;
    const kAnt = t === 'D' ? diaMas(fecha, -1) : fecha, tAnt = t === 'D' ? 'N' : 'D';
    const ant = resumen(kAnt === fecha ? doc : await datos.dia(kAnt), kAnt, tAnt, cfg);
    asunto = `Planta Aerobol · Turno ${r.nombre.toLowerCase()} ${fechaCorta(fecha)} · ${nf0.format(r.total)} envases${r.meta ? ` (${nf0.format(Math.floor(r.pct))} % de la meta)` : ''}`;
    html = correo(`Resumen del turno ${r.nombre.toLowerCase()}`, `${fechaLarga(fecha)} · ${r.horario}`, bloqueTurno(r, ant, false));
  }

  if (programado && totalEnviado === 0 && String(process.env.REPORTE_SI_CERO || '').trim().toLowerCase() !== 'si') {
    console.log('El turno no tiene producción registrada: no se envía el correo (REPORTE_SI_CERO distinto de "si").');
    return;
  }
  if (process.env.DRY_RUN === '1') { fs.writeFileSync('reporte.html', html); console.log('Asunto:', asunto, '\nGuardado en reporte.html (no se envió).'); return; }
  const prueba = String(process.env.SOLO_PRUEBA) === 'true';
  const para = prueba ? (process.env.REPORTE_PRUEBA || process.env.REPORTE_PARA) : process.env.REPORTE_PARA;
  if (!para) throw new Error('Falta REPORTE_PARA (destinatarios) en las variables del repositorio.');
  await enviar({ asunto: (prueba ? '[PRUEBA] ' : '') + asunto, html, para, cc: prueba ? '' : process.env.REPORTE_CC });
  console.log('Enviado:', asunto, '→', para);
}
main().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
