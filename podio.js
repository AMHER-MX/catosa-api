// ── RUTA AL PODIO · Lógica de puntos (Sep 1 – Nov 30, 2026) ──────────────────
// Módulo puro (sin SQL ni red) para poder probarlo aislado. server.js lo alimenta.
'use strict';

const PODIO = {
  INI: '2026-09-01',
  FIN: '2026-11-30',
  MESES: ['2026-09', '2026-10', '2026-11'],
  CANALES: ['CALLE', 'MOSTRADOR'],
  PTS: {
    CHECKIN_DIA:      10,   // ambos canales
    CONSTANCIA_DIA:   10,   // día que supera el promedio diario de meta
    RACHA_DIA_EXTRA:  10,   // extra por cada día consecutivo (a partir del 2º)
    VISITAS_SEMANA:   25,   // CALLE: ≥15 visitas/semana
    VISITAS_MIN:      15,
    ENCUESTA_SEMANA:  25,   // MOSTRADOR: ≥15 escaneos/semana
    ENCUESTA_MIN:     15,
    FLEETRITE: [ { min: 50, pts: 100 }, { min: 40, pts: 75 }, { min: 30, pts: 50 } ],
  },
  TEMAS: {
    '2026-09': { nombre: 'Viva México',            emoji: '🇲🇽' },
    '2026-10': { nombre: 'Halloween',              emoji: '🎃' },
    '2026-11': { nombre: 'Día de Muertos + Navidad', emoji: '💀🎄' },
  },
  PUESTOS: ['Oro', 'Plata', 'Bronce', '4to', '5to'],
};

// ── Utilidades de fecha (todas en strings YYYY-MM-DD, sin zonas horarias) ────
function pad(n) { return String(n).padStart(2, '0'); }
function isoDe(d) { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
function fechaLocal(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function hoyISO() { return isoDe(new Date()); }
function diaSemana(iso) { const [y, m, d] = iso.split('-').map(Number); return new Date(y, m - 1, d).getDay(); } // 0=Dom
function ultimoDiaMes(mes) { const [y, m] = mes.split('-').map(Number); return new Date(y, m, 0).getDate(); }
function mesDe(iso) { return (iso || '').slice(0, 7); }
function sumarDias(iso, n) { const [y, m, d] = iso.split('-').map(Number); return isoDe(new Date(y, m - 1, d + n)); }
function lunesDe(iso) { const dow = diaSemana(iso); return sumarDias(iso, dow === 0 ? -6 : 1 - dow); }

// Convierte 'dd/mm/aaaa', 'd/m/aa', 'aaaa-mm-dd' o Date → 'YYYY-MM-DD' (o null)
function normalizarFecha(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : isoDe(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${pad(m[2])}-${pad(m[1])}`; }
  return null;
}

// Días hábiles (lun–sáb) de un mes
function diasHabilesMes(mes) {
  let n = 0;
  for (let d = 1; d <= ultimoDiaMes(mes); d++) if (diaSemana(`${mes}-${pad(d)}`) !== 0) n++;
  return n;
}

// Mes del concurso al que pertenece una fecha, o null si está fuera
function mesConcurso(iso) {
  if (!iso || iso < PODIO.INI || iso > PODIO.FIN) return null;
  return mesDe(iso);
}

function temaDe(mes) { return PODIO.TEMAS[mes] || { nombre: 'Ruta al Podio', emoji: '🏁' }; }

// ── CSV mínimo (maneja comillas y comas dentro de comillas) ───────────────────
function parseCSV(text) {
  const rows = []; let row = [], cell = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(cell); cell = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += c;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  if (!rows.length) return [];
  const head = rows[0].map(h => h.trim());
  return rows.slice(1).filter(r => r.some(x => x && x.trim())).map(r => {
    const o = {}; head.forEach((h, i) => { o[h] = (r[i] || '').trim(); }); return o;
  });
}
// Busca una columna sin importar mayúsculas/acentos
function col(row, ...nombres) {
  const keys = Object.keys(row);
  for (const n of nombres) {
    const k = keys.find(k => k.toLowerCase() === n.toLowerCase());
    if (k !== undefined) return row[k];
  }
  return '';
}
function nombreKey(s) { return (s || '').toString().trim().toUpperCase().replace(/\s+/g, ' '); }

// ── Reglas ───────────────────────────────────────────────────────────────────

// Constancia + racha. ventasDia: { 'YYYY-MM-DD': monto }. Solo días lun–sáb ≤ hoy.
function calcConstancia(ventasDia, meta, mes, hoy) {
  const habiles = diasHabilesMes(mes);
  const metaDiaria = meta > 0 && habiles > 0 ? meta / habiles : 0;
  const dias = [];
  let racha = 0, rachaMax = 0, diasSuperados = 0, ptsBase = 0, ptsRacha = 0;
  const ult = ultimoDiaMes(mes);
  for (let d = 1; d <= ult; d++) {
    const f = `${mes}-${pad(d)}`;
    if (f > hoy) break;
    if (diaSemana(f) === 0) continue;                  // domingo: no cuenta ni rompe racha
    const venta = parseFloat(ventasDia[f]) || 0;
    const superado = metaDiaria > 0 && venta > metaDiaria;
    if (superado) {
      diasSuperados++; ptsBase += PODIO.PTS.CONSTANCIA_DIA;
      racha++;
      if (racha >= 2) ptsRacha += PODIO.PTS.RACHA_DIA_EXTRA;
      rachaMax = Math.max(rachaMax, racha);
    } else if (f < hoy) {
      racha = 0;                                        // día hábil pasado sin superar: rompe
    }
    // si f === hoy y aún no supera, la racha sigue "viva" hasta que cierre el día
    dias.push({ fecha: f, venta: Math.round(venta), superado });
  }
  return { metaDiaria: Math.round(metaDiaria), diasHabiles: habiles, diasSuperados,
           rachaActual: racha, rachaMax, ptsBase, ptsRacha, pts: ptsBase + ptsRacha, dias };
}

// Check-in: 1 vez por día (fechas ISO, se deduplican)
function calcCheckin(fechas, mes) {
  const set = new Set((fechas || []).filter(f => f && mesDe(f) === mes));
  return { dias: set.size, pts: set.size * PODIO.PTS.CHECKIN_DIA, fechas: [...set].sort() };
}

// Semanas (lunes a domingo) con ≥ min registros → pts por semana.
// Una semana pertenece al mes en que cae su lunes.
function calcSemanal(fechas, mes, min, ptsSemana) {
  const porSemana = {};
  (fechas || []).forEach(f => {
    if (!f) return;
    const lun = lunesDe(f);
    if (mesDe(lun) !== mes) return;
    porSemana[lun] = (porSemana[lun] || 0) + 1;
  });
  // incluir todas las semanas del mes (para mostrar avance aunque tenga 0)
  let lun = lunesDe(`${mes}-01`);
  if (mesDe(lun) !== mes) lun = sumarDias(lun, 7);
  const semanas = [];
  const hoy = hoyISO();
  for (; mesDe(lun) === mes; lun = sumarDias(lun, 7)) {
    if (lun > hoy) break;
    const total = porSemana[lun] || 0;
    semanas.push({ inicio: lun, fin: sumarDias(lun, 6), total, cumple: total >= min });
  }
  const cumplidas = semanas.filter(s => s.cumple).length;
  return { semanas, cumplidas, pts: cumplidas * ptsSemana, meta: min };
}

function calcFleetrite(nps) {
  const n = parseInt(nps) || 0;
  const nivel = PODIO.PTS.FLEETRITE.find(t => n >= t.min);
  const siguiente = [...PODIO.PTS.FLEETRITE].reverse().find(t => n < t.min);
  return { nps: n, pts: nivel ? nivel.pts : 0, nivel: nivel ? nivel.min : 0,
           siguiente: siguiente ? siguiente.min : null, faltan: siguiente ? siguiente.min - n : 0 };
}

// ── Cálculo de un mes ────────────────────────────────────────────────────────
// datos = {
//   mes, hoy,
//   participantes: [{ Nombre, Sucursal, Canal, Meta }],
//   ventasDia:  { NOMBRE: { 'YYYY-MM-DD': monto } },
//   checkins:   { NOMBRE: ['YYYY-MM-DD', ...] },
//   visitas:    { NOMBRE: ['YYYY-MM-DD', ...] },
//   encuestas:  { NOMBRE: ['YYYY-MM-DD', ...] },
//   fleetrite:  { NOMBRE: nps_distintos },
//   ventasMes:  { NOMBRE: monto }
// }
function calcularMes(datos) {
  const { mes } = datos;
  const hoy = datos.hoy || hoyISO();
  return datos.participantes.map(p => {
    const k = nombreKey(p.Nombre);
    const esCalle = p.Canal === 'CALLE';
    const checkin    = calcCheckin(datos.checkins[k], mes);
    const constancia = calcConstancia(datos.ventasDia[k] || {}, p.Meta, mes, hoy);
    const visitas    = esCalle ? calcSemanal(datos.visitas[k], mes, PODIO.PTS.VISITAS_MIN, PODIO.PTS.VISITAS_SEMANA) : null;
    const encuesta   = !esCalle ? calcSemanal(datos.encuestas[k], mes, PODIO.PTS.ENCUESTA_MIN, PODIO.PTS.ENCUESTA_SEMANA) : null;
    const fleetrite  = calcFleetrite(datos.fleetrite[k]);
    const total = checkin.pts + constancia.pts + (visitas ? visitas.pts : 0) + (encuesta ? encuesta.pts : 0) + fleetrite.pts;
    return {
      Nombre: p.Nombre, Sucursal: p.Sucursal, Canal: p.Canal, Meta: p.Meta,
      Ventas: Math.round(datos.ventasMes[k] || 0),
      Pct: p.Meta > 0 ? Math.round(((datos.ventasMes[k] || 0) / p.Meta) * 1000) / 10 : 0,
      checkin, constancia, visitas, encuesta, fleetrite, total,
    };
  });
}

// Ordena y asigna puesto. Desempate: constancia, luego % de meta, luego nombre.
function rankear(lista, campoTotal = 'total') {
  const orden = [...lista].sort((a, b) =>
    (b[campoTotal] - a[campoTotal]) ||
    ((b.constanciaPts || 0) - (a.constanciaPts || 0)) ||
    ((b.Pct || 0) - (a.Pct || 0)) ||
    a.Nombre.localeCompare(b.Nombre));
  orden.forEach((r, i) => { r.puesto = i + 1; r.medalla = PODIO.PUESTOS[i] || null; });
  return orden;
}

// Combina los meses calculados en un acumulado del concurso
function acumular(porMes, mesActual) {
  const mapa = {};
  Object.entries(porMes).forEach(([mes, lista]) => {
    lista.forEach(r => {
      const k = nombreKey(r.Nombre);
      if (!mapa[k]) mapa[k] = { Nombre: r.Nombre, Sucursal: r.Sucursal, Canal: r.Canal, Meta: r.Meta,
                                total: 0, constanciaPts: 0, meses: {}, Ventas: 0, Pct: 0 };
      const m = mapa[k];
      m.total += r.total; m.constanciaPts += r.constancia.pts;
      m.meses[mes] = { total: r.total, checkin: r.checkin.pts, constancia: r.constancia.pts,
                       visitas: r.visitas ? r.visitas.pts : 0, encuesta: r.encuesta ? r.encuesta.pts : 0,
                       fleetrite: r.fleetrite.pts, Ventas: r.Ventas, Pct: r.Pct };
      if (mes === mesActual) { m.Ventas = r.Ventas; m.Pct = r.Pct; m.actual = r; }
    });
  });
  return rankear(Object.values(mapa));
}

module.exports = { PODIO, parseCSV, col, nombreKey, normalizarFecha, diasHabilesMes, mesConcurso, temaDe,
                   calcConstancia, calcCheckin, calcSemanal, calcFleetrite, calcularMes, rankear, acumular,
                   hoyISO, lunesDe, sumarDias, ultimoDiaMes, pad };
