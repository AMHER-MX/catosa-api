require('dotenv').config();
const express = require('express');
const cors    = require('cors');
const sql     = require('mssql');
const XLSX    = require('xlsx');
const path    = require('path');

const app = express();
app.use(cors());
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

// ── CARGA DEL EXCEL DE METAS ─────────────────────────────────────────────────
let metasMap   = {};   // { "VICTOR CASTILLO": { meta, canal, sucursal } }
let carteraMap = {};   // { "VICTOR CASTILLO": ["CLIENTE A", ...] }

function cargarExcel() {
  try {
    const wb  = XLSX.readFile(path.join(__dirname, 'metas.xlsx'));

    // Hoja OBJETIVOS VENTA
    const obj = XLSX.utils.sheet_to_json(wb.Sheets['OBJETIVOS VENTA'], { defval: null });
    obj.forEach(row => {
      const nombre = (row['NOMBRE ASESOR'] || '').toString().trim().toUpperCase();
      const meta   = parseFloat(row['META MENSUAL']) || 0;
      const canal  = (row['CANAL']    || '').toString().trim().toUpperCase();
      const suc    = (row['SUCURSAL'] || '').toString().trim().toUpperCase();
      if (nombre && meta > 0) metasMap[nombre] = { meta, canal, sucursal: suc };
    });

    // Hoja CARTERA
    const car = XLSX.utils.sheet_to_json(wb.Sheets['CARTERA'], { defval: null });
    car.forEach(row => {
      const v = (row['Vendedor '] || row['Vendedor'] || '').toString().trim().toUpperCase();
      const c = (row['Cliente ']  || row['Cliente']  || '').toString().trim().toUpperCase();
      if (!v || !c) return;
      if (!carteraMap[v]) carteraMap[v] = [];
      carteraMap[v].push(c);
    });

    console.log(`Excel cargado: ${Object.keys(metasMap).length} asesores, ${Object.values(carteraMap).flat().length} clientes`);
  } catch (err) {
    console.error('Error leyendo metas.xlsx:', err.message);
  }
}

cargarExcel();

// Recarga Excel sin reiniciar: GET /api/recargar-metas
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

function buscarCartera(vendedor) {
  const vKey = nombreKey(vendedor);
  for (const [k, v] of Object.entries(carteraMap)) {
    if (k.includes(vKey) || vKey.includes(k.split(' ')[0])) return v;
  }
  return [];
}

// ── VENTAS + METAS ───────────────────────────────────────────────────────────
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
          v.NOM_VENDEDOR AS Nombre,
          v.ALMACEN      AS Sucursal_SQL,
          COALESCE(SUM(p.P_VENTA * p.CANTIDAD_PEDIDA), 0) AS Ventas
        FROM FMVENBI_PR v
        LEFT JOIN FTPDCBI_PR p
          ON  p.ALMACEN = v.ALMACEN
          AND p.FECHA  >= @ini
          AND p.FECHA  <= @fin
          AND p.PEDIDO_CANCELADO <> 'S'
        WHERE v.EN_LISTA = 'S'
        GROUP BY v.NOM_VENDEDOR, v.ALMACEN
        ORDER BY Ventas DESC
      `);

    const datos = result.recordset.map(row => {
      const m = buscarMeta(row.Nombre);
      return {
        Nombre:     row.Nombre,
        Sucursal:   m.sucursal || row.Sucursal_SQL,
        Canal:      m.canal,
        Ventas:     row.Ventas,
        Meta:       m.meta,
        Venta_Prov: 0,
        Meta_Prov:  50000,
      };
    });

    res.json(datos);
  } catch (err) {
    console.error('Error /api/ventas:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── CLIENTES DEL VENDEDOR ────────────────────────────────────────────────────
app.get('/api/clientes', async (req, res) => {
  try {
    const db     = await getPool();
    const hoy    = new Date();
    const ini    = `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,'0')}-01`;
    const hoyStr = hoy.toISOString().split('T')[0];

    const result = await db.request()
      .input('ini', sql.Date, ini)
      .input('hoy', sql.Date, hoyStr)
      .query(`
        SELECT
          p.CTA_CLIENTE                                      AS Cliente,
          DATEDIFF(DAY, MAX(p.FECHA), @hoy)                 AS Dias,
          SUM(CASE WHEN p.FECHA >= @ini
                   THEN p.P_VENTA * p.CANTIDAD_PEDIDA ELSE 0
              END)                                           AS Venta_Mes
        FROM FTPDCBI_PR p
        WHERE p.PEDIDO_CANCELADO <> 'S'
        GROUP BY p.CTA_CLIENTE
        ORDER BY Dias ASC
      `);

    let clientes = result.recordset;

    // Filtra por cartera del Excel si hay match
    const cartera = buscarCartera(req.query.vendedor || '');
    if (cartera.length > 0) {
      clientes = clientes.filter(c =>
        cartera.some(cf => nombreKey(c.Cliente).includes(cf.split(' ')[0]) || cf.includes(nombreKey(c.Cliente).split(' ')[0]))
      );
    }

    res.json(clientes);
  } catch (err) {
    console.error('Error /api/clientes:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── INVENTARIO ───────────────────────────────────────────────────────────────
app.get('/api/productos', async (req, res) => {
  try {
    const db     = await getPool();
    const { sku, q } = req.query;
    const busq   = sku || q || '';
    const exacto = !!sku;

    const result = await db.request()
      .input('b', sql.VarChar, exacto ? busq : `%${busq}%`)
      .query(`
        SELECT TOP 20
          i.ARTICULO     AS Parte,
          i.DES_ARTICULO AS Descripcion,
          i.EXIS_REALES  AS Existencia,
          i.COSTO_MEDIO  AS Precio,
          i.UBICACION    AS Ubicacion
        FROM FTIGBI_PR i
        WHERE ${exacto ? 'i.ARTICULO = @b' : 'i.ARTICULO LIKE @b OR i.DES_ARTICULO LIKE @b'}
        ORDER BY i.EXIS_REALES DESC
      `);

    res.json(result.recordset);
  } catch (err) {
    console.error('Error /api/productos:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── TOP 10 PRODUCTOS ─────────────────────────────────────────────────────────
app.get('/api/top-productos', async (req, res) => {
  try {
    const db  = await getPool();
    const hoy = new Date();
    const ini = `${hoy.getFullYear()}-${String(hoy.getMonth()+1).padStart(2,'0')}-01`;
    const cartera = buscarCartera(req.query.vendedor || '');

    const result = await db.request()
      .input('ini', sql.Date, ini)
      .query(`
        SELECT TOP 10
          p.ARTICULO                 AS Parte,
          p.DES_ARTICULO             AS Descripcion,
          SUM(p.CANTIDAD_PEDIDA)     AS Cantidad,
          COALESCE(i.EXIS_REALES, 0) AS Existencia
        FROM FTPDCBI_PR p
        LEFT JOIN FTIGBI_PR i ON i.ARTICULO = p.ARTICULO
        WHERE p.FECHA >= @ini
          AND p.PEDIDO_CANCELADO <> 'S'
          ${cartera.length > 0 ? `AND p.CTA_CLIENTE IN (${cartera.map(c => `'${c.replace(/'/g,"''")}'`).join(',')})` : ''}
        GROUP BY p.ARTICULO, p.DES_ARTICULO, i.EXIS_REALES
        ORDER BY Cantidad DESC
      `);

    res.json(result.recordset);
  } catch (err) {
    console.error('Error /api/top-productos:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Health check simple — NO requiere DB
app.get('/', (req, res) => res.json({ status: 'ok' }));

// ── PING ─────────────────────────────────────────────────────────────────────
app.get('/api/ping', async (req, res) => {
  try {
    const db = await getPool();
    await db.request().query('SELECT 1');
    res.json({ status: 'ok', db: 'conectado', asesores: Object.keys(metasMap).length });
  } catch (err) {
    res.status(500).json({ status: 'error', mensaje: err.message });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`Catosa API en http://localhost:${PORT}`));
