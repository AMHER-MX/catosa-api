require('dotenv').config({ path: 'C:\\catosa-api\\.env' });
const express = require('express');
const cors    = require('cors');
const sql     = require('mssql');
const XLSX    = require('xlsx');
const path    = require('path');

const app = express();
app.use(cors({ origin: '*', methods: ['GET','POST','OPTIONS'], allowedHeaders: ['Content-Type','Authorization'] }));
app.use(express.json());

const dbConfig = {
  server:   process.env.DB_HOST,
  user:     process.env.DB_USER,
  password: process.env.DB_PASS,
  database: process.env.DB_NAME,
  port:     parseInt(process.env.DB_PORT) || 1433,
  options:  { encrypt: false, trustServerCertificate: true },
  connectionTimeout: 15000,
  requestTimeout:    30000,
};

let pool;
async function getPool() {
  if (!pool) pool = await sql.connect(dbConfig);
  return pool;
}

// ── EXCEL DE METAS ────────────────────────────────────────────────────────────
let metasMap = {}, carteraMap = {}, aceiteBaseMap = {};

function cargarExcel() {
  try {
    const wb  = XLSX.readFile(path.join(__dirname, 'metas.xlsx'));
    const obj = XLSX.utils.sheet_to_json(wb.Sheets['OBJETIVOS VENTA'], { defval: null });
    obj.forEach(row => {
      const nombre = (row['NOMBRE ASESOR'] || '').toString().trim().toUpperCase();
      const meta   = parseFloat(row['META MENSUAL']) || 0;
      const canal  = (row['CANAL']    || '').toString().trim().toUpperCase();
      const suc    = (row['SUCURSAL'] || '').toString().trim().toUpperCase();
      if (nombre && meta > 0) metasMap[nombre] = { meta, canal, sucursal: suc };
    });
    const car = XLSX.utils.sheet_to_json(wb.Sheets['CARTERA'], { defval: null });
    car.forEach(row => {
      const v = (row['Vendedor '] || row['Vendedor'] || '').toString().trim().toUpperCase();
      const c = (row['Cliente ']  || row['Cliente']  || '').toString().trim().toUpperCase();
      if (!v || !c) return;
      if (!carteraMap[v]) carteraMap[v] = [];
      carteraMap[v].push(c);
    });
    // Carga promedios base de aceite 2025
    const aceite = XLSX.utils.sheet_to_json(wb.Sheets['ACEITE_BASE'], { defval: null });
    aceite.forEach(row => {
      const nombre = (row['NOMBRE_ASESOR'] || '').toString().trim().toUpperCase();
      const prom   = parseFloat(row['PROMEDIO_LITROS_2025']) || 0;
      if (nombre) aceiteBaseMap[nombre] = prom;
    });
    console.log(`Excel cargado: ${Object.keys(metasMap).length} asesores, ${Object.values(carteraMap).flat().length} clientes, ${Object.keys(aceiteBaseMap).length} bases de aceite`);
  } catch (err) { console.error('Error leyendo metas.xlsx:', err.message); }
}
cargarExcel();

app.get('/api/recargar-metas', (req, res) => {
  metasMap = {}; carteraMap = {};
  cargarExcel();
  res.json({ ok: true, asesores: Object.keys(metasMap).length });
});

function nombreKey(s) { return (s || '').toString().trim().toUpperCase(); }

function buscarMeta(nombreSql) {
  const key = nombreKey(nombreSql);
  if (metasMap[key]) return metasMap[key];
  for (const [k, v] of Object.entries(metasMap)) {
    if (key.includes(k.split(' ')[0]) || k.includes(key.split(' ')[0])) return v;
  }
  return { meta: 0, canal: 'CALLE', sucursal: '' };
}

// Busca cartera con match flexible (igual que buscarMeta)
function buscarCartera(vendedor) {
  const key = nombreKey(vendedor);
  if (carteraMap[key]) return carteraMap[key];
  for (const [k, v] of Object.entries(carteraMap)) {
    if (key.includes(k.split(' ')[0]) || k.includes(key.split(' ')[0])) return v;
  }
  return [];
}

// Metas por sucursal (definidas por dirección)
const META_SUCURSAL = {
  'TORREON':        5000000,
  'GOMEZ PALACIO':  7000000,
  'MONCLOVA':       2100000,
  'PIEDRAS NEGRAS': 2100000,
  'ANA':            0,
};

const SUCURSAL_NORM = {
  'TR': 'TORREON', 'TORREÓN': 'TORREON',
  'GP': 'GOMEZ PALACIO', 'GÓMEZ PALACIO': 'GOMEZ PALACIO',
  'MONC': 'MONCLOVA', 'PN': 'PIEDRAS NEGRAS',
};
function normSuc(s) { const k = (s||'').toUpperCase().trim(); return SUCURSAL_NORM[k] || k; }

const SUCURSALES    = `'ANA','GOMEZ PALACIO','MONCLOVA','PIEDRAS NEGRAS','TORREON'`;
const TIPOS_EXCL    = `'PRESUPUESTO','PRESUPUESTO 8%','Traspaso salida almacen'`;
const TIPO_EXCL_SQL = `(s.DES_TIPO_VENTA NOT IN (${TIPOS_EXCL}) AND s.DES_TIPO_VENTA IS NOT NULL AND LTRIM(RTRIM(s.DES_TIPO_VENTA)) <> '')`;

// ── HEALTH CHECK ──────────────────────────────────────────────────────────────
app.get('/', (req, res) => res.json({ status: 'ok', servidor: 'Catosa API' }));

app.get('/api/ping', async (req, res) => {
  try {
    const db = await getPool();
    await db.request().query('SELECT 1');
    res.json({ status: 'ok', db: 'conectado', asesores: Object.keys(metasMap).length });
  } catch (err) { res.status(500).json({ status: 'error', mensaje: err.message }); }
});

// ── VENTAS + METAS ────────────────────────────────────────────────────────────
// ── VENTAS TOTALES REALES (todos los vendedores, para KPIs de suma) ───────────
app.get('/api/ventas-totales', async (req, res) => {
  try {
    const db  = await getPool();
    const hoy = new Date();
    const ini = `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,'0')}-01`;
    const fin = hoy.toISOString().split('T')[0];

    const result = await db.request()
      .input('ini', sql.Date, ini)
      .input('fin', sql.Date, fin)
      .query(`
        SELECT
          s.NOM_ALMACEN_LIN            AS Sucursal,
          SUM(s.IMP_TOTAL_LINEA)       AS Ventas
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
        GROUP BY s.NOM_ALMACEN_LIN
        ORDER BY Ventas DESC
      `);

    // Total global
    const total = result.recordset.reduce((sum, r) => sum + (parseFloat(r.Ventas)||0), 0);
    res.json({ total, porSucursal: result.recordset });
  } catch (err) {
    console.error('Error /api/ventas-totales:', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/ventas', async (req, res) => {
  try {
    const db  = await getPool();
    const hoy = new Date();
    const ini = `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,'0')}-01`;
    const fin = hoy.toISOString().split('T')[0];

    const result = await db.request()
      .input('ini', sql.Date, ini)
      .input('fin', sql.Date, fin)
      .query(`
        SELECT
          s.NOM_VENDEDOR             AS Nombre,
          s.NOM_ALMACEN_LIN          AS Sucursal_SQL,
          SUM(s.IMP_TOTAL_LINEA)     AS Ventas
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR IS NOT NULL AND s.NOM_VENDEDOR <> ''
        GROUP BY s.NOM_VENDEDOR, s.NOM_ALMACEN_LIN
        ORDER BY Ventas DESC
      `);

    const datos = result.recordset.map(row => {
      const m = buscarMeta(row.Nombre);
      const metaSuc = META_SUCURSAL[normSuc(m.sucursal || row.Sucursal_SQL)] || 0;
      return {
        Nombre: row.Nombre, Sucursal: normSuc(m.sucursal || row.Sucursal_SQL),
        Canal: m.canal, Ventas: row.Ventas, Meta: m.meta,
        MetaSucursal: metaSuc,
        Venta_Prov: 0, Meta_Prov: 50000,
      };
    });
    res.json(datos);
  } catch (err) {
    console.error('Error /api/ventas:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── CLIENTES DEL VENDEDOR (con nombre comercial y dirección) ──────────────────
app.get('/api/clientes', async (req, res) => {
  try {
    const db      = await getPool();
    const hoy     = new Date();
    const ini     = `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,'0')}-01`;
    const hoyStr  = hoy.toISOString().split('T')[0];
    const vendedor = decodeURIComponent(req.query.vendedor || '');

    const result = await db.request()
      .input('ini',  sql.Date, ini)
      .input('hoy',  sql.Date, hoyStr)
      .input('vend', sql.VarChar, `%${vendedor}%`)
      .query(`
        SELECT
          s.CLIENTE                                                          AS Codigo,
          COALESCE(NULLIF(c.NOMBRE_COMERCIAL,''), c.NOMBRE, s.CLIENTE)      AS Cliente,
          COALESCE(c.NOMBRE, '')                                             AS NombreCompleto,
          DATEDIFF(DAY, MAX(s.FECHA), @hoy)                                 AS Dias,
          SUM(CASE WHEN s.FECHA >= @ini THEN s.IMP_TOTAL_LINEA ELSE 0 END)  AS Venta_Mes,
          SUM(CASE WHEN s.FECHA >= DATEADD(year,-1,@ini) AND s.FECHA < @ini
              THEN s.IMP_TOTAL_LINEA ELSE 0 END) / 12.0                     AS PromedioMensual,
          COALESCE(c.DIRECCION, '')                                          AS Direccion,
          COALESCE(c.DES_COLONIA, '')                                        AS Colonia,
          COALESCE(c.DES_DELEGACION, '')                                     AS Ciudad,
          COALESCE(c.DES_REGION, '')                                         AS Estado,
          COALESCE(c.COD_POSTAL, '')                                         AS CP,
          COALESCE(c.PAIS, '')                                               AS Pais
        FROM FTSABI_PR s
        LEFT JOIN FMCUBI_PR c ON c.CUENTA = s.CLIENTE
        WHERE ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR LIKE @vend
        GROUP BY s.CLIENTE, c.NOMBRE_COMERCIAL, c.NOMBRE,
                 c.DIRECCION, c.DES_COLONIA, c.DES_DELEGACION,
                 c.DES_REGION, c.COD_POSTAL, c.PAIS
        ORDER BY Dias ASC
      `);

    // Armar dirección + marcar si es de cartera
    const cartera    = buscarCartera(vendedor);
    const carteraSet = new Set(cartera.map(c => c.toUpperCase().trim()));

    const datos = result.recordset.map(r => ({
      ...r,
      Direccion: [r.Direccion, r.Colonia, r.Ciudad, r.Estado, r.CP]
        .filter(Boolean).join(', '),
      EsCartera: carteraSet.has((r.Codigo || '').toUpperCase().trim()),
      PromedioMensual: parseFloat(r.PromedioMensual) || 0
    }));

    res.json(datos);
  } catch (err) {
    console.error('Error /api/clientes:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── INVENTARIO / COTIZADOR ────────────────────────────────────────────────────
app.get('/api/productos', async (req, res) => {
  try {
    const db  = await getPool();
    const { sku, q } = req.query;
    const busq  = sku || q || '';
    const exacto = !!sku;

    const result = await db.request()
      .input('b', sql.VarChar, exacto ? busq : `%${busq}%`)
      .query(`
        SELECT TOP 20
          i.ARTICULO      AS Parte,
          i.DES_ARTICULO  AS Descripcion,
          (SELECT SUM(i2.EXIS_REALES) FROM FTIGBI_PR i2 
           WHERE i2.ARTICULO = i.ARTICULO AND i2.ALMACEN IN ('101', '102', '101LA', '102LA')) AS Existencia,
          i.COSTO_MEDIO   AS Precio,
          i.UBICACION     AS Ubicacion
        FROM FTIGBI_PR i
        WHERE i.ALMACEN = 101
          AND ${exacto
          ? 'i.ARTICULO = @b'
          : 'i.ARTICULO LIKE @b OR i.DES_ARTICULO LIKE @b'}
        ORDER BY Existencia DESC
      `);

    res.json(result.recordset);
  } catch (err) {
    console.error('Error /api/productos:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── TOP 10 PRODUCTOS DEL VENDEDOR ─────────────────────────────────────────────
app.get('/api/top-productos', async (req, res) => {
  try {
    const db      = await getPool();
    const hoy     = new Date();
    const ini     = `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,'0')}-01`;
    const vendedor = decodeURIComponent(req.query.vendedor || '');

    const result = await db.request()
      .input('ini',  sql.Date, ini)
      .input('vend', sql.VarChar, `%${vendedor}%`)
      .query(`
        SELECT TOP 10
          s.ARTICULO                              AS Parte,
          s.DES_ARTICULO                          AS Descripcion,
          SUM(s.CANTIDAD)                         AS Unidades_Vendidas,
          SUM(s.IMP_TOTAL_LINEA)                  AS Monto,
          COALESCE(inv.Existencia, 0)             AS Existencia
        FROM FTSABI_PR s
        LEFT JOIN (
          SELECT ARTICULO, SUM(EXIS_REALES) AS Existencia
          FROM FTIGBI_PR
          WHERE ALMACEN IN ('101', '102', '101LA', '102LA')
          GROUP BY ARTICULO
        ) inv ON inv.ARTICULO = s.ARTICULO
        WHERE s.FECHA >= @ini
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR LIKE @vend
        GROUP BY s.ARTICULO, s.DES_ARTICULO, inv.Existencia
        ORDER BY Unidades_Vendidas DESC
      `);

    res.json(result.recordset);
  } catch (err) {
    console.error('Error /api/top-productos:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── CHALLENGE ACEITE MOTOR ────────────────────────────────────────────────────
// Números de parte participantes con prefijos 0/ y 1/ como aparecen en FTSABI_PR
const ACEITE_NPS = {};
const _NPS_BASE = {
  // Terminación BK = 1 litro
  'FLRTASCI4PLUSBK':1,'INTLCI4PBLK':1,'FLRTCI4BK':1,'FLRT15W40CK4BK':1,'INTLCK4BK':1,
  'INT15W40CI4PBLK':1,'INT15W40CK4BLK':1,'FLRT3460BK':1,'FLRTCK4BK':1,
  // Terminación DR = 208 litros
  'FLRTASCI4PLUSDR':208,'INTLCI4PDR':208,'FLRT15W40CI4DR':208,'FLRT15W40CK4DR':208,'INTLCK4DR':208,
  'INT15W40CI4PDR':208,'INT15W40CK4DR':208,'FLRT3460DR':208,'FLRTCK4DR':208,'INT71231328D':208,'FLRT25W50DR':208,
  // Terminación PL = 19 litros
  'FLRTASCI4PLUSPL':19,'INTLCI4PPL':19,'FLRT15W40CI4PL':19,'FLRT15W40CK4PL':19,'INTLCK4PL':19,
  'INT15W40CI4PPL':19,'INT15W40CK4PL':19,'FLRT3460PL':19,'FLRTCK4PL':19,'INT71231319P':19,'FLRT25W50PL':19,
  // Terminación TL = 1000 litros
  'FLRTASCI4PLUSTL':1000,'INTLCI4PTL':1000,'FLRT15W40CI4TL':1000,'FLRT15W40CK4TL':1000,'INTLCK4TL':1000,
  'INT15W40CI4PTT':1000,'INT15W40CK4TT':1000,'FLRT3460TL':1000,'FLRTCK4TL':1000,
  // Terminación GA = 4 litros
  'FLRT15W40CI4G':4,'FLRT3460GA':4,'INT71231305G':4,
};
// Agrega prefijos 0/ y 1/ a cada NP
for (const [np, lts] of Object.entries(_NPS_BASE)) {
  ACEITE_NPS[np]       = lts;
  ACEITE_NPS['0/'+np]  = lts;
  ACEITE_NPS['1/'+np]  = lts;
}
// Números de marca International (con prefijos)
const _INTL_BASE = ['INTLCI4PBLK','INTLCI4PDR','INTLCI4PPL','INTLCI4PTL',
  'INTLCK4BK','INTLCK4DR','INTLCK4PL','INTLCK4TL',
  'INT15W40CI4PBLK','INT15W40CI4PDR','INT15W40CI4PPL','INT15W40CI4PTT',
  'INT15W40CK4BLK','INT15W40CK4DR','INT15W40CK4PL','INT15W40CK4TT',
  'INT71231305G','INT71231319P','INT71231328D'];
const NPS_INTL = new Set([..._INTL_BASE, ..._INTL_BASE.map(n=>'0/'+n), ..._INTL_BASE.map(n=>'1/'+n)]);

app.get('/api/aceite', async (req, res) => {
  try {
    const db      = await getPool();
    const hoy     = new Date();
    const ini     = '2026-05-01'; // Concurso acumulativo mayo-diciembre 2026
    const fin     = hoy.toISOString().split('T')[0];
    const vendedor = decodeURIComponent(req.query.vendedor || '');

    // Todos los NPs participantes
    const npsLista = Object.keys(ACEITE_NPS).map(n => `'${n}'`).join(',');

    const result = await db.request()
      .input('ini',  sql.Date, ini)
      .input('fin',  sql.Date, fin)
      .input('vend', sql.VarChar, vendedor)
      .query(`
        SELECT s.ARTICULO, SUM(s.CANTIDAD) AS Cantidad
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND s.ARTICULO IN (${npsLista})
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR = @vend
          AND ${TIPO_EXCL_SQL}
        GROUP BY s.ARTICULO
      `);

    // Calcular litros totales e International
    let litrosTotal = 0, litrosIntl = 0;
    result.recordset.forEach(row => {
      const factor  = ACEITE_NPS[row.ARTICULO] || 1;
      const litros  = (parseFloat(row.Cantidad) || 0) * factor;
      litrosTotal  += litros;
      if (NPS_INTL.has(row.ARTICULO)) litrosIntl += litros;
    });

    // Base 2025 del vendedor — match exacto con nombres SQL
    const vKey = nombreKey(vendedor);
    const basePromedio = aceiteBaseMap[vKey] || 0;

    const litrosIncrementales = Math.max(0, litrosTotal - basePromedio);
    const bloques             = Math.floor(litrosIncrementales / 500);
    const pctIntl             = litrosTotal > 0 ? (litrosIntl / litrosTotal) * 100 : 0;
    const premioPorBloque     = pctIntl >= 30 ? 500 : 400;
    const premioTotal         = bloques * premioPorBloque;

    res.json({
      litrosTotal:        Math.round(litrosTotal),
      litrosIntl:         Math.round(litrosIntl),
      litrosIncrementales:Math.round(litrosIncrementales),
      basePromedio:       Math.round(basePromedio),
      bloques,
      pctIntl:            pctIntl.toFixed(1),
      premioPorBloque,
      premioTotal,
      // Para la barra: % de avance hacia el siguiente bloque de 500L
      pctBarra: Math.min(100, ((litrosIncrementales % 500) / 500) * 100),
      bloquesSiguiente: bloques + 1,
    });
  } catch (err) {
    console.error('Error /api/aceite:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── ACEITE TODOS LOS VENDEDORES (para gerencia) ───────────────────────────────
app.get('/api/aceite-todos', async (req, res) => {
  try {
    const db  = await getPool();
    const hoy = new Date();
    const ini = '2026-05-01'; // Concurso acumulativo mayo-diciembre 2026
    const fin = hoy.toISOString().split('T')[0];
    const npsLista = Object.keys(ACEITE_NPS).map(n => `'${n}'`).join(',');

    const result = await db.request()
      .input('ini', sql.Date, ini)
      .input('fin', sql.Date, fin)
      .query(`
        SELECT s.NOM_VENDEDOR, s.ARTICULO, SUM(s.CANTIDAD) AS Cantidad
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND s.ARTICULO IN (${npsLista})
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND ${TIPO_EXCL_SQL}
        GROUP BY s.NOM_VENDEDOR, s.ARTICULO
      `);

    // Agrupar por vendedor
    const porVendedor = {};
    result.recordset.forEach(row => {
      const v = row.NOM_VENDEDOR;
      if (!porVendedor[v]) porVendedor[v] = { litrosTotal: 0, litrosIntl: 0 };
      const factor = ACEITE_NPS[row.ARTICULO] || 1;
      const litros = (parseFloat(row.Cantidad) || 0) * factor;
      porVendedor[v].litrosTotal += litros;
      if (NPS_INTL.has(row.ARTICULO)) porVendedor[v].litrosIntl += litros;
    });

    // Construir respuesta para cada vendedor con meta
    const datos = Object.entries(metasMap).map(([nombre, datos]) => {
      const nombreSQL = Object.keys(porVendedor).find(k => nombreKey(k) === nombre) || '';
      const vData = porVendedor[nombreSQL] || { litrosTotal: 0, litrosIntl: 0 };
      const base  = aceiteBaseMap[nombre] || 0;
      const litrosIncrementales = Math.max(0, vData.litrosTotal - base);
      const bloques = Math.floor(litrosIncrementales / 500);
      const pctIntl = vData.litrosTotal > 0 ? (vData.litrosIntl / vData.litrosTotal) * 100 : 0;
      const premioPorBloque = pctIntl >= 30 ? 500 : 400;
      return {
        Nombre:              nombreSQL || nombre,
        Sucursal:            datos.sucursal,
        LitrosTotal:         Math.round(vData.litrosTotal),
        LitrosIncrementales: Math.round(litrosIncrementales),
        BasePromedio:        Math.round(base),
        Bloques:             bloques,
        PctIntl:             pctIntl.toFixed(1),
        PremioPorBloque:     premioPorBloque,
        PremioTotal:         bloques * premioPorBloque,
        PctBarra:            Math.min(100, ((litrosIncrementales % 500) / 500) * 100),
      };
    }).sort((a, b) => b.LitrosIncrementales - a.LitrosIncrementales);

    res.json(datos);
  } catch (err) {
    console.error('Error /api/aceite-todos:', err.message);
    res.status(500).json({ error: err.message });
  }
});
app.get('/api/ventas-diarias', async (req, res) => {
  try {
    const db      = await getPool();
    const hoy     = new Date();
    const ini     = '2026-05-01'; // Concurso acumulativo mayo-diciembre 2026
    const fin     = hoy.toISOString().split('T')[0];
    const vendedor = decodeURIComponent(req.query.vendedor || '');

    const result = await db.request()
      .input('ini',  sql.Date, ini)
      .input('fin',  sql.Date, fin)
      .input('vend', sql.VarChar, vendedor)
      .query(`
        SELECT
          DAY(s.FECHA)               AS Dia,
          SUM(s.IMP_TOTAL_LINEA)     AS Venta
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR = @vend
        GROUP BY DAY(s.FECHA)
        ORDER BY Dia ASC
      `);

    // Devuelve objeto { 1: 1234.5, 2: 0, 3: 890, ... } para los 31 días
    const dias = {};
    result.recordset.forEach(r => { dias[r.Dia] = parseFloat(r.Venta) || 0; });
    res.json(dias);
  } catch (err) {
    console.error('Error /api/ventas-diarias:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── CORES PENDIENTES DE PAGO (FTPDCBI_PR sin FEC_CANCELACION) ─────────────────
app.get('/api/cores-pendientes', async (req, res) => {
  try {
    const db       = await getPool();
    const vendedor = decodeURIComponent(req.query.vendedor || '');
    const cartera  = buscarCartera(vendedor);

    if (cartera.length === 0) return res.json([]);

    const clientesIn = cartera.map(c => `'${c.replace(/'/g,"''")}'`).join(',');

    const result = await db.request().query(`
      SELECT
        p.REFERENCIA                              AS Referencia,
        p.NRO_PEDIDO                              AS NroPedido,
        p.CTA_CLIENTE                             AS Codigo,
        p.DES_ARTICULO                            AS Articulo,
        p.DES_TIPO_VENTA                          AS TipoVenta,
        CONVERT(varchar(10), p.FECHA, 23)         AS FechaFactura,
        p.P_VENTA                                 AS Monto,
        p.CANTIDAD_PEDIDA                         AS CantidadPedida,
        p.CANTIDAD_RECIBIDA                       AS CantidadRecibida,
        DATEDIFF(day, p.FECHA, GETDATE())         AS DiasSinPagar
      FROM FTPDCBI_PR p
      WHERE p.DES_TIPO_VENTA IN ('VENTA REMISIONES CORES', 'VENTA REMISIONES CORES 8%')
        AND p.FEC_CANCELACION IS NULL
        AND p.CTA_CLIENTE IN (${clientesIn})
      ORDER BY DiasSinPagar DESC
    `);

    res.json(result.recordset);
  } catch (err) {
    console.error('Error /api/cores-pendientes:', err.message);
    res.status(500).json({ error: err.message });
  }
});
app.get('/api/cores', async (req, res) => {
  try {
    const db       = await getPool();
    const vendedor = decodeURIComponent(req.query.vendedor || '');
    const hoy      = new Date();
    const ini      = `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,'0')}-01`;
    const fin      = hoy.toISOString().split('T')[0];

    const result = await db.request()
      .input('ini',  sql.Date, ini)
      .input('fin',  sql.Date, fin)
      .input('vend', sql.VarChar, `%${vendedor}%`)
      .query(`
        SELECT
          s.REFERENCIA                          AS Referencia,
          s.NUM_FACTURA                         AS NumFactura,
          s.CLIENTE                             AS Codigo,
          s.NOM_VENDEDOR                        AS Vendedor,
          s.DES_ARTICULO                        AS Articulo,
          s.DES_TIPO_VENTA                      AS TipoVenta,
          CONVERT(varchar(10), s.FECHA, 23)     AS FechaFactura,
          s.IMP_TOTAL_LINEA                     AS Monto,
          s.CANTIDAD                            AS Cantidad,
          DATEDIFF(day, s.FECHA, GETDATE())     AS DiasTranscurridos
        FROM FTSABI_PR s
        WHERE s.DES_TIPO_VENTA IN (
            'CANCELACION VENTA REMISIONES CORES',
            'CANCELACION VENTA REMISIONES CORES 8%'
          )
          AND s.FECHA >= @ini AND s.FECHA <= @fin
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR LIKE @vend
        ORDER BY s.FECHA DESC
      `);

    res.json(result.recordset);
  } catch (err) {
    console.error('Error /api/cores:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── TORNEO WORLD CUP — Clasificación real desde SQL ──────────────────────────
app.get('/api/torneo', async (req, res) => {
  try {
    const db   = await getPool();
    const anio = new Date().getFullYear();

    // Ventas Mayo 18–31
    const rMayo = await db.request().query(`
      SELECT NOM_VENDEDOR, SUM(IMP_TOTAL_LINEA) AS VentaMayo
      FROM FTSABI_PR
      WHERE FECHA >= '${anio}-05-18' AND FECHA <= '${anio}-05-31'
        AND DES_TIPO_VENTA NOT IN (${TIPOS_EXCL}) AND DES_TIPO_VENTA IS NOT NULL AND LTRIM(RTRIM(DES_TIPO_VENTA)) <> ''
        AND NOM_ALMACEN_LIN IN (${SUCURSALES})
      GROUP BY NOM_VENDEDOR
    `);
    const ventaMayo = {};
    rMayo.recordset.forEach(r => { ventaMayo[nombreKey(r.NOM_VENDEDOR)] = parseFloat(r.VentaMayo)||0; });

    // Ventas Junio completo
    const rJunio = await db.request().query(`
      SELECT NOM_VENDEDOR, SUM(IMP_TOTAL_LINEA) AS VentaJunio
      FROM FTSABI_PR
      WHERE FECHA >= '${anio}-06-01' AND FECHA <= '${anio}-06-30'
        AND DES_TIPO_VENTA NOT IN (${TIPOS_EXCL}) AND DES_TIPO_VENTA IS NOT NULL AND LTRIM(RTRIM(DES_TIPO_VENTA)) <> ''
        AND NOM_ALMACEN_LIN IN (${SUCURSALES})
      GROUP BY NOM_VENDEDOR
    `);
    const ventaJunio = {};
    rJunio.recordset.forEach(r => { ventaJunio[nombreKey(r.NOM_VENDEDOR)] = parseFloat(r.VentaJunio)||0; });

    // Ventas Julio 1–17 (fase final)
    const rJulio = await db.request().query(`
      SELECT NOM_VENDEDOR, SUM(IMP_TOTAL_LINEA) AS VentaJulio
      FROM FTSABI_PR
      WHERE FECHA >= '${anio}-07-01' AND FECHA <= '${anio}-07-17'
        AND DES_TIPO_VENTA NOT IN (${TIPOS_EXCL}) AND DES_TIPO_VENTA IS NOT NULL AND LTRIM(RTRIM(DES_TIPO_VENTA)) <> ''
        AND NOM_ALMACEN_LIN IN (${SUCURSALES})
      GROUP BY NOM_VENDEDOR
    `);
    const ventaJulio = {};
    rJulio.recordset.forEach(r => { ventaJulio[nombreKey(r.NOM_VENDEDOR)] = parseFloat(r.VentaJulio)||0; });

    // Armar resultado por vendedor
    const resultado = Object.entries(metasMap).map(([nombre, datos]) => {
      const meta     = datos.meta || 1;
      const vMayo    = ventaMayo[nombre]  || 0;
      const vJunio   = ventaJunio[nombre] || 0;
      const vJulio   = ventaJulio[nombre] || 0;
      const pctMayo  = (vMayo  / meta) * 100;
      const pctJunio = (vJunio / meta) * 100;
      const pctJulio = (vJulio / meta) * 100;
      const clasificado = pctMayo >= 120 || pctJunio >= 120;
      return {
        Nombre:       nombre,
        Sucursal:     datos.sucursal,
        Meta:         meta,
        VentaMayo:    vMayo,
        VentaJunio:   vJunio,
        VentaJulio:   vJulio,
        PctMayo:      parseFloat(pctMayo.toFixed(2)),
        PctJunio:     parseFloat(pctJunio.toFixed(2)),
        PctJulio:     parseFloat(pctJulio.toFixed(2)),
        Clasificado:  clasificado,
        ClasifMayo:   pctMayo  >= 120,
        ClasifJunio:  pctJunio >= 120,
      };
    });

    res.json(resultado.sort((a, b) => (b.Clasificado - a.Clasificado) || b.PctMayo - a.PctMayo));
  } catch (err) {
    console.error('Error /api/torneo:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── TABLAS: exponer tablas de SQL Server para app de tableros ──────────────
app.get('/api/tables', async (req, res) => {
  try {
    const db = await getPool();
    const result = await db.request().query(
      "SELECT name FROM sys.tables ORDER BY name"
    );
    res.json({ tables: result.recordset.map(r => r.name) });
  } catch (err) {
    pool = null;
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/table/:name', async (req, res) => {
  const tableName = req.params.name.replace(/[^a-zA-Z0-9_]/g, '');
  const limit = Math.min(parseInt(req.query.limit) || 500, 50000);
  const dateFrom = req.query.dateFrom;
  const dateTo   = req.query.dateTo;
  const dateCol  = req.query.dateCol;
  try {
    const db = await getPool();
    const colRes = await db.request().query(
      `SELECT name FROM sys.columns WHERE object_id = OBJECT_ID('${tableName}') ORDER BY column_id`
    );
    const columns = colRes.recordset.map(r => r.name);

    let where = '';
    if (dateCol && columns.includes(dateCol)) {
      const parts = [];
      if (dateFrom) parts.push(`[${dateCol}] >= '${dateFrom}'`);
      if (dateTo)   parts.push(`[${dateCol}] <= '${dateTo} 23:59:59'`);
      if (parts.length) where = 'WHERE ' + parts.join(' AND ');
    }

    const data = await db.request().query(
      `SELECT TOP ${limit} * FROM [${tableName}] ${where}`
    );
    res.json({ columns, rows: data.recordset, total: data.recordset.length });
  } catch (err) {
    pool = null;
    res.status(500).json({ error: err.message });
  }
});

// ── TABLEROS: guardar/cargar por área ─────────────────────────────────────
const fs = require('fs');
const TABLEROS_FILE = path.join(__dirname, 'tableros.json');

function leerTableros() {
  try { return JSON.parse(fs.readFileSync(TABLEROS_FILE, 'utf8')); }
  catch { return {}; }
}
function guardarTableros(data) {
  fs.writeFileSync(TABLEROS_FILE, JSON.stringify(data, null, 2));
}

app.get('/api/tableros/:area', (req, res) => {
  const data = leerTableros();
  res.json(data[req.params.area] || []);
});

app.post('/api/tableros/:area', (req, res) => {
  const data = leerTableros();
  data[req.params.area] = req.body;
  guardarTableros(data);
  res.json({ ok: true });
});

app.delete('/api/tableros/:area/:id', (req, res) => {
  const data = leerTableros();
  if (data[req.params.area]) {
    data[req.params.area] = data[req.params.area].filter(d => d.id !== req.params.id);
    guardarTableros(data);
  }
  res.json({ ok: true });
});


// ── RESUMEN MES ANTERIOR ───────────────────────────────────────────────────────
app.get('/api/resumen-mes', async (req, res) => {
  try {
    const db      = await getPool();
    const hoy     = new Date();
    const vendedor = decodeURIComponent(req.query.vendedor || '');

    // Mes anterior
    const iniAnt  = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
    const finAnt  = new Date(hoy.getFullYear(), hoy.getMonth(), 0);
    const iniStr  = iniAnt.toISOString().split('T')[0];
    const finStr  = finAnt.toISOString().split('T')[0];
    const mesNombre = iniAnt.toLocaleString('es-MX', { month: 'long', year: 'numeric' });

    // ── 1. Ventas del mes anterior (individual) ────────────────────────────
    const rVentas = await db.request()
      .input('ini',  sql.Date, iniStr)
      .input('fin',  sql.Date, finStr)
      .input('vend', sql.VarChar, vendedor)
      .query(`
        SELECT SUM(s.IMP_TOTAL_LINEA) AS Ventas
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR = @vend
      `);
    const ventasAnt = parseFloat(rVentas.recordset[0]?.Ventas) || 0;

    // ── 2. Meta del vendedor ───────────────────────────────────────────────
    const vKey  = nombreKey(vendedor);
    const metaD = Object.entries(metasMap).find(([k]) => k === vKey);
    const meta  = metaD ? metaD[1].meta : 0;
    const pct   = meta > 0 ? (ventasAnt / meta) * 100 : 0;

    // ── 3. Top 5 productos mes anterior ───────────────────────────────────
    const rProd = await db.request()
      .input('ini',  sql.Date, iniStr)
      .input('fin',  sql.Date, finStr)
      .input('vend', sql.VarChar, `%${vendedor}%`)
      .query(`
        SELECT TOP 5
          s.ARTICULO AS Parte, s.DES_ARTICULO AS Descripcion,
          SUM(s.CANTIDAD) AS Unidades, SUM(s.IMP_TOTAL_LINEA) AS Monto
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR LIKE @vend
        GROUP BY s.ARTICULO, s.DES_ARTICULO
        ORDER BY Monto DESC
      `);

    // ── 4. Top 5 clientes mes anterior ────────────────────────────────────
    const rCli = await db.request()
      .input('ini',  sql.Date, iniStr)
      .input('fin',  sql.Date, finStr)
      .input('vend', sql.VarChar, `%${vendedor}%`)
      .query(`
        SELECT TOP 5
          s.CLIENTE AS Codigo,
          MAX(s.NOMBRE_CLIENTE) AS Cliente,
          SUM(s.IMP_TOTAL_LINEA) AS Monto
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR LIKE @vend
        GROUP BY s.CLIENTE
        ORDER BY Monto DESC
      `);

    // ── 5. Días trabajados (días con al menos 1 venta) ────────────────────
    const rDias = await db.request()
      .input('ini',  sql.Date, iniStr)
      .input('fin',  sql.Date, finStr)
      .input('vend', sql.VarChar, `%${vendedor}%`)
      .query(`
        SELECT COUNT(DISTINCT CAST(s.FECHA AS DATE)) AS Dias
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR LIKE @vend
      `);
    const diasTrabajados = parseInt(rDias.recordset[0]?.Dias) || 0;
    const diasHabiles = 22; // promedio

    // ── 6. Ventas mes anterior del equipo (para gerencia) ─────────────────
    const rEquipo = await db.request()
      .input('ini', sql.Date, iniStr)
      .input('fin', sql.Date, finStr)
      .query(`
        SELECT
          s.NOM_VENDEDOR AS Nombre,
          s.NOM_ALMACEN_LIN AS Sucursal_SQL,
          SUM(s.IMP_TOTAL_LINEA) AS Ventas
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.NOM_VENDEDOR IS NOT NULL AND s.NOM_VENDEDOR <> ''
        GROUP BY s.NOM_VENDEDOR, s.NOM_ALMACEN_LIN
        ORDER BY Ventas DESC
      `);

    const equipo = rEquipo.recordset.map(row => {
      const m = buscarMeta(row.Nombre);
      const metaV = m.meta || 0;
      const pctV  = metaV > 0 ? (parseFloat(row.Ventas) / metaV) * 100 : 0;
      return {
        Nombre:   row.Nombre,
        Sucursal: normSuc(m.sucursal || row.Sucursal_SQL),
        Ventas:   parseFloat(row.Ventas) || 0,
        Meta:     metaV,
        Pct:      Math.round(pctV),
        Cumple:   pctV >= 100,
      };
    });

    const totalEquipo    = equipo.reduce((s, v) => s + v.Ventas, 0);
    const metaEquipo     = equipo.reduce((s, v) => s + v.Meta, 0);
    const pctEquipo      = metaEquipo > 0 ? (totalEquipo / metaEquipo) * 100 : 0;
    const cumplieron     = equipo.filter(v => v.Cumple).length;

    // ── Áreas de oportunidad ──────────────────────────────────────────────
    const oportunidades = [];
    if (pct < 100) {
      const faltante = meta - ventasAnt;
      oportunidades.push(`Faltaron ${new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',minimumFractionDigits:0}).format(faltante)} para llegar a meta (${Math.round(pct)}% alcanzado)`);
    }
    if (diasTrabajados < diasHabiles * 0.8) {
      oportunidades.push(`Solo se registraron ventas en ${diasTrabajados} días — mayor cobertura diaria puede impulsar el resultado`);
    }
    if (pct >= 100) {
      oportunidades.push(`¡Meta cumplida! Mantén el ritmo y busca superar el ${Math.round(pct)}% este mes`);
    }
    if (rProd.recordset.length > 0) {
      const top1 = rProd.recordset[0];
      oportunidades.push(`Tu producto estrella fue ${top1.Descripcion || top1.Parte} — considera ampliar su penetración en más clientes`);
    }

    res.json({
      mes: mesNombre,
      vendedor,
      ventasAnt,
      meta,
      pct: Math.round(pct),
      cumpleMeta: pct >= 100,
      diasTrabajados,
      topProductos: rProd.recordset,
      topClientes:  rCli.recordset,
      oportunidades,
      // Para gerencia
      equipo,
      totalEquipo,
      metaEquipo,
      pctEquipo: Math.round(pctEquipo),
      cumplieron,
      totalAsesores: equipo.length,
    });
  } catch (err) {
    console.error('Error /api/resumen-mes:', err.message);
    res.status(500).json({ error: err.message });
  }
});


// ── DETALLE CLIENTE ───────────────────────────────────────────────────────────
app.get('/api/cliente-detalle', async (req, res) => {
  try {
    const db      = await getPool();
    const hoy     = new Date();
    const hoyStr  = hoy.toISOString().split('T')[0];
    const hace1a  = new Date(hoy); hace1a.setFullYear(hace1a.getFullYear()-1);
    const hace6m  = new Date(hoy); hace6m.setMonth(hace6m.getMonth()-6);
    const hace1aStr = hace1a.toISOString().split('T')[0];
    const hace6mStr = hace6m.toISOString().split('T')[0];
    const cliente  = decodeURIComponent(req.query.cliente || '');
    const vendedor = decodeURIComponent(req.query.vendedor || '');

    // ── Top 5 productos ───────────────────────────────────────────────────
    const rProd = await db.request()
      .input('cli',   sql.VarChar, cliente)
      .input('vend',  sql.VarChar, `%${vendedor}%`)
      .input('ini1a', sql.VarChar, hace1aStr)
      .query(`
        SELECT TOP 5
          s.ARTICULO                        AS Parte,
          s.DES_ARTICULO                    AS Descripcion,
          SUM(s.CANTIDAD)                   AS Unidades,
          SUM(s.IMP_TOTAL_LINEA)            AS Monto
        FROM FTSABI_PR s
        WHERE s.CLIENTE = @cli
          AND s.NOM_VENDEDOR LIKE @vend
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.FECHA >= @ini1a
        GROUP BY s.ARTICULO, s.DES_ARTICULO
        ORDER BY Monto DESC
      `);

    // ── Ventas por mes (últimos 6 meses) ──────────────────────────────────
    const rMeses = await db.request()
      .input('cli',   sql.VarChar, cliente)
      .input('vend',  sql.VarChar, `%${vendedor}%`)
      .input('ini6m', sql.VarChar, hace6mStr)
      .query(`
        SELECT
          LEFT(s.FECHA, 7)               AS Mes,
          SUM(s.IMP_TOTAL_LINEA)         AS Venta
        FROM FTSABI_PR s
        WHERE s.CLIENTE = @cli
          AND s.NOM_VENDEDOR LIKE @vend
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND s.FECHA >= @ini6m
        GROUP BY LEFT(s.FECHA, 7)
        ORDER BY Mes ASC
      `);

    // ── Última compra ─────────────────────────────────────────────────────
    const rUlt = await db.request()
      .input('cli',  sql.VarChar, cliente)
      .input('vend', sql.VarChar, `%${vendedor}%`)
      .query(`
        SELECT TOP 1
          s.FECHA               AS Fecha,
          s.DES_ARTICULO        AS Producto,
          s.IMP_TOTAL_LINEA     AS Monto
        FROM FTSABI_PR s
        WHERE s.CLIENTE = @cli
          AND s.NOM_VENDEDOR LIKE @vend
          AND ${TIPO_EXCL_SQL}
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
        ORDER BY s.FECHA DESC
      `);

    res.json({
      topProductos: rProd.recordset,
      meses:        rMeses.recordset,
      ultimaCompra: rUlt.recordset[0] || null,
    });
  } catch (err) {
    console.error('Error /api/cliente-detalle:', err.message);
    res.status(500).json({ error: err.message });
  }
});


// ── REPORTE ACEITE POR MES (para Excel) ──────────────────────────────────────
app.get('/api/aceite-reporte', async (req, res) => {
  try {
    const db  = await getPool();
    const mes = req.query.mes || ''; // formato YYYY-MM
    if (!mes) return res.status(400).json({ error: 'Parámetro mes requerido (YYYY-MM)' });

    const ini = `${mes}-01`;
    // Último día del mes
    const [y, m] = mes.split('-').map(Number);
    const fin = new Date(y, m, 0).toISOString().split('T')[0];

    const npsLista = Object.keys(ACEITE_NPS).map(n => `'${n}'`).join(',');
    const mesNombre = new Date(y, m-1, 1).toLocaleString('es-MX', { month: 'long', year: 'numeric' })
                        .toUpperCase();

    const result = await db.request()
      .input('ini', sql.VarChar, ini)
      .input('fin', sql.VarChar, fin)
      .query(`
        SELECT
          s.NOM_VENDEDOR  AS Vendedor,
          s.ARTICULO      AS NumeroParte,
          SUM(s.CANTIDAD) AS Cantidad
        FROM FTSABI_PR s
        WHERE s.FECHA >= @ini AND s.FECHA <= @fin
          AND s.ARTICULO IN (${npsLista})
          AND s.NOM_ALMACEN_LIN IN (${SUCURSALES})
          AND ${TIPO_EXCL_SQL}
        GROUP BY s.NOM_VENDEDOR, s.ARTICULO
        ORDER BY s.NOM_VENDEDOR
      `);

    // Agrupar por vendedor
    const porVendedor = {};
    result.recordset.forEach(row => {
      const v = row.Vendedor;
      if (!porVendedor[v]) porVendedor[v] = { partes: {}, litrosTotal: 0, litrosIntl: 0, litrosFlrt: 0 };
      const factor = ACEITE_NPS[row.NumeroParte] || 1;
      const litros = (parseFloat(row.Cantidad) || 0) * factor;
      porVendedor[v].litrosTotal += litros;
      porVendedor[v].partes[row.NumeroParte] = (porVendedor[v].partes[row.NumeroParte] || 0) + litros;
      if (NPS_INTL.has(row.NumeroParte)) porVendedor[v].litrosIntl += litros;
      else porVendedor[v].litrosFlrt += litros;
    });

    // Construir filas del reporte
    const filas = Object.entries(porVendedor).map(([vendedor, data]) => {
      const npsVendidos = Object.entries(data.partes)
        .sort((a, b) => b[1] - a[1])
        .map(([np, lts]) => `${np} (${Math.round(lts)}L)`)
        .join(', ');
      return {
        Distribuidor:           'CATOSA INTERNATIONAL',
        GFX:                    'zz001',
        Vendedor:               vendedor,
        Mes_Venta:              mesNombre,
        Volumen_Venta:          Math.round(data.litrosTotal),
        Litros_International:   Math.round(data.litrosIntl),
        Litros_Fleetrite:       Math.round(data.litrosFlrt),
        Numeros_Parte_Vendidos: npsVendidos,
        Litros_Por_Taller:      0,
      };
    }).sort((a, b) => b.Volumen_Venta - a.Volumen_Venta);

    res.json({ mes: mesNombre, filas });
  } catch (err) {
    console.error('Error /api/aceite-reporte:', err.message);
    res.status(500).json({ error: err.message });
  }
});


const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => console.log(`Catosa API en http://localhost:${PORT}`));
