const express = require('express');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execSync } = require('child_process');
process.env.PUPPETEER_CACHE_DIR = path.join(__dirname, '.cache', 'puppeteer');
const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const session = require('express-session');
const passport = require('passport');
const GoogleStrategy = require('passport-google-oauth20').Strategy;
const app = express();
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '0.0.0.0';
const APP_BASE_URL = process.env.APP_BASE_URL || `http://localhost:${PORT}`;
const APP_BASE_URL_NORMALIZED = String(APP_BASE_URL).replace(/\/+$/, '');
const SESSION_SECRET = process.env.SESSION_SECRET || 'dev-session-secret-change-me';
const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '';
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || '';
const AUTH_READY = Boolean(GOOGLE_CLIENT_ID && GOOGLE_CLIENT_SECRET);
const ALLOWED_DOMAIN = String(process.env.ALLOWED_DOMAIN || 'simca.mx').toLowerCase();
const GERENTE_EMAIL = String(process.env.GERENTE_EMAIL || 'martin@simca.mx').toLowerCase();
const LOG_PATH = '/tmp/fr-ven-server.log';
const log = (msg) => {
  const line = `[${new Date().toISOString()}] ${msg}\n`;
  try { fs.appendFileSync(LOG_PATH, line); } catch {}
  console.log(msg);
};

log('Iniciando servidor...');
const PUPPETEER_CACHE_DIR = process.env.PUPPETEER_CACHE_DIR || path.join(__dirname, '.cache', 'puppeteer');

const TEMPLATE_DIR = path.join(__dirname, 'templates');
const PUBLIC_DIR = path.join(__dirname, 'public');
const DATA_PATH = path.join(__dirname, 'data', 'sample.json');
const DATA_DIR = path.join(__dirname, 'data');
const ROI_MASTER_CSV_PATH = path.join(DATA_DIR, 'roi-master.csv');
const SUBMISSIONS_PATH = path.join(DATA_DIR, 'submissions.json');
const OWNER_SERVICES_PATH = path.join(DATA_DIR, 'owner-services.json');
const FLOOR_JSON_DIR = path.join(DATA_DIR, 'plano-ventas-floors');
const DEVELOPMENTS_DIR = path.join(DATA_DIR, 'developments');
const DEFAULT_DEVELOPMENT_SLUG = 'ceiba';
const FLOOR_JSON_FILE_RE = /^unidades-marcadas(?:\s*\(\d+\))?\.json$/i;
const DEVELOPMENTS = [
  { slug: 'ceiba', name: 'CEIBA' },
  { slug: 'costa-caribe', name: 'COSTA CARIBE' },
  { slug: 'cruz-con-mar', name: 'CRUZ CON MAR' },
  { slug: 'dream-c', name: 'DREAM C' },
  { slug: 'gran-tulum', name: 'GRAN TULUM' },
  { slug: 'ipana', name: 'IPANA' },
  { slug: 'maresol', name: 'MARESOL' },
  { slug: 'marila', name: 'MARILA' },
  { slug: 'natal', name: 'NATAL' },
  { slug: 'saint-marine', name: 'SAINT MARINE' },
  { slug: 'serenada', name: 'SERENADA' },
  { slug: 'singular-joy', name: 'SINGULAR JOY' },
  { slug: 'solar', name: 'SOLAR' },
  { slug: 'solar-mt', name: 'SOLAR MT' }
];
const DEVELOPMENTS_BY_SLUG = DEVELOPMENTS.reduce((acc, item) => {
  acc[item.slug] = item;
  return acc;
}, {});

const formats = {
  '35': { name: 'FR-VEN-35 Aviso de Privacidad', file: 'format-35.html' },
  '21': { name: 'FR-VEN-21 Consulta de Listas de Personas Bloqueadas', file: 'format-21.html' },
  '26': { name: 'FR-VEN-26 Origen de los Recursos', file: 'format-26.html' },
  '19': { name: 'FR-VEN-19 Beneficiario Controlador PF', file: 'format-19.html' },
  '10': { name: 'FR-VEN-10 Identificación del Cliente (Extranjera PF)', file: 'format-10.html' }
};

app.set('trust proxy', 1);
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(session({
  name: 'simca.sid',
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    sameSite: 'lax',
    secure: process.env.NODE_ENV === 'production',
    maxAge: 1000 * 60 * 60 * 12
  }
}));
app.use(passport.initialize());
app.use(passport.session());
app.use('/assets', express.static(path.join(PUBLIC_DIR, 'assets')));
app.use('/plds-static', express.static(path.join(PUBLIC_DIR, 'plds')));

passport.serializeUser((user, done) => done(null, user));
passport.deserializeUser((user, done) => done(null, user));

if (AUTH_READY) {
  passport.use(new GoogleStrategy(
    {
      clientID: GOOGLE_CLIENT_ID,
      clientSecret: GOOGLE_CLIENT_SECRET,
      callbackURL: `${APP_BASE_URL_NORMALIZED}/auth/google/callback`
    },
    (accessToken, refreshToken, profile, done) => {
      const rawEmail = profile && profile.emails && profile.emails[0] ? profile.emails[0].value : '';
      const email = String(rawEmail || '').trim().toLowerCase();
      if (!email || !email.endsWith(`@${ALLOWED_DOMAIN}`)) {
        return done(null, false, { message: `Solo cuentas @${ALLOWED_DOMAIN}` });
      }
      return done(null, {
        id: profile.id,
        email,
        name: profile.displayName || email
      });
    }
  ));
}

function requireAuth(req, res, next) {
  if (req.isAuthenticated && req.isAuthenticated()) return next();
  if (String(req.path || '').startsWith('/api/')) {
    return res.status(401).json({ error: 'No autenticado' });
  }
  return res.redirect('/login');
}

function requireGerente(req, res, next) {
  if (!(req.isAuthenticated && req.isAuthenticated())) {
    if (String(req.path || '').startsWith('/api/')) {
      return res.status(401).json({ error: 'No autenticado' });
    }
    return res.redirect('/login');
  }
  const email = String(req.user && req.user.email || '').toLowerCase();
  if (email === GERENTE_EMAIL) return next();
  if (String(req.path || '').startsWith('/api/')) {
    return res.status(403).json({ error: 'No autorizado para este módulo' });
  }
  return res.status(403).send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Acceso restringido</title>
  <style>
    :root{--bg:#f4f1e8;--card:#fff;--line:#d8d1c1;--ink:#1a1a1a;--muted:#5f5f5f;}
    *{box-sizing:border-box}
    body{margin:0;font-family:Arial,sans-serif;background:var(--bg);color:var(--ink);display:grid;place-items:center;min-height:100vh;padding:20px;}
    .card{width:min(620px,100%);background:var(--card);border:1px solid var(--line);border-radius:14px;padding:22px;}
    h1{margin:0 0 10px;font-size:28px;}
    p{margin:0 0 12px;color:var(--muted);line-height:1.4;}
    .btn{display:inline-block;padding:10px 12px;border:1px solid #bdb8a9;border-radius:10px;background:#fff;color:#111;text-decoration:none;font-size:13px;font-weight:600;}
  </style></head>
  <body>
    <section class="card">
      <h1>Acceso restringido</h1>
      <p>Esta sección es exclusiva para el usuario gerente.</p>
      <p>Usuario autorizado: <strong>${GERENTE_EMAIL}</strong>.</p>
      <a class="btn" href="/">Regresar al inicio</a>
    </section>
  </body></html>`);
}

process.on('uncaughtException', (err) => {
  log(`Uncaught exception: ${err && err.stack ? err.stack : err}`);
});
process.on('unhandledRejection', (err) => {
  log(`Unhandled rejection: ${err && err.stack ? err.stack : err}`);
});

function ensureDataFiles() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(FLOOR_JSON_DIR)) fs.mkdirSync(FLOOR_JSON_DIR, { recursive: true });
  if (!fs.existsSync(DEVELOPMENTS_DIR)) fs.mkdirSync(DEVELOPMENTS_DIR, { recursive: true });
  DEVELOPMENTS.forEach((dev) => {
    const floorDir = path.join(DEVELOPMENTS_DIR, dev.slug, 'plano-ventas-floors');
    if (!fs.existsSync(floorDir)) fs.mkdirSync(floorDir, { recursive: true });
  });
  if (!fs.existsSync(SUBMISSIONS_PATH)) fs.writeFileSync(SUBMISSIONS_PATH, '[]', 'utf-8');
  if (!fs.existsSync(DATA_PATH)) fs.writeFileSync(DATA_PATH, '{}', 'utf-8');
  if (!fs.existsSync(OWNER_SERVICES_PATH)) {
    const initialOwnerServices = {
      project: {
        name: 'Proyecto principal',
        regimenCondominial: false,
        regimenDate: '',
        internalFinancingReady: false,
        updatedAt: new Date().toISOString()
      },
      units: []
    };
    fs.writeFileSync(OWNER_SERVICES_PATH, JSON.stringify(initialOwnerServices, null, 2), 'utf-8');
  }
}

function normalizeDevelopmentSlug(raw) {
  const value = String(raw || '').trim().toLowerCase();
  if (!value) return DEFAULT_DEVELOPMENT_SLUG;
  if (DEVELOPMENTS_BY_SLUG[value]) return value;
  return DEFAULT_DEVELOPMENT_SLUG;
}

function getDevelopmentBySlug(raw) {
  const slug = normalizeDevelopmentSlug(raw);
  return DEVELOPMENTS_BY_SLUG[slug] || DEVELOPMENTS_BY_SLUG[DEFAULT_DEVELOPMENT_SLUG];
}

function getRequestedDevelopment(req) {
  const slugFromPath = req && req.params ? req.params.devSlug : '';
  const slugFromQuery = req && req.query ? req.query.dev : '';
  return getDevelopmentBySlug(slugFromPath || slugFromQuery || DEFAULT_DEVELOPMENT_SLUG);
}

function getDevelopmentFloorDir(devSlug) {
  return path.join(DEVELOPMENTS_DIR, devSlug, 'plano-ventas-floors');
}

function getDevelopmentFloorSearchDirs(devSlug) {
  const primary = getDevelopmentFloorDir(devSlug);
  if (devSlug === DEFAULT_DEVELOPMENT_SLUG) {
    return [primary, FLOOR_JSON_DIR];
  }
  return [primary];
}

function listFloorJsonFiles(dir) {
  try {
    return fs.readdirSync(dir)
      .filter((name) => FLOOR_JSON_FILE_RE.test(name))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
  } catch {
    return [];
  }
}

function sanitizeJsonFileName(rawName) {
  const fallback = 'unidades-marcadas-unificado-orden-adjuntos.json';
  const value = String(rawName || '').trim();
  if (!value) return fallback;
  const base = path.basename(value).replace(/[^\w\-(). ]+/g, '_');
  const normalized = base.toLowerCase().endsWith('.json') ? base : `${base}.json`;
  return normalized || fallback;
}

function launchPdfBrowser() {
  const launchOptions = {
    headless: 'new',
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--no-zygote',
      '--single-process',
      '--font-render-hinting=none'
    ]
  };
  return puppeteer.launch(launchOptions);
}

function installChromeIfMissing(err) {
  const msg = String((err && err.message) || '');
  if (!/Could not find Chrome/i.test(msg)) return false;
  try {
    log(`Chrome no encontrado. Instalando en caliente en ${PUPPETEER_CACHE_DIR} ...`);
    execSync('npx puppeteer browsers install chrome', {
      stdio: 'inherit',
      env: {
        ...process.env,
        PUPPETEER_CACHE_DIR
      }
    });
    return true;
  } catch (installErr) {
    log(`Fallo instalando Chrome en caliente: ${installErr && installErr.stack ? installErr.stack : installErr}`);
    return false;
  }
}

function sanitizeExcelFileName(rawName) {
  const value = String(rawName || '').trim();
  if (!value) return '';
  const base = path.basename(value).replace(/[^\w\-(). ]+/g, '_');
  const ext = (base.split('.').pop() || '').toLowerCase();
  if (!['xls', 'xlsx', 'csv'].includes(ext)) return '';
  return base;
}

function readJson(filePath, fallback) {
  try {
    if (!fs.existsSync(filePath)) return fallback;
    return JSON.parse(fs.readFileSync(filePath, 'utf-8'));
  } catch (err) {
    log(`Error leyendo ${filePath}: ${err.message}`);
    return fallback;
  }
}

function writeJson(filePath, data) {
  const tmpPath = `${filePath}.tmp`;
  fs.writeFileSync(tmpPath, JSON.stringify(data, null, 2), 'utf-8');
  fs.renameSync(tmpPath, filePath);
}

function ownerServicesDefaultData() {
  return {
    project: {
      name: 'Proyecto principal',
      regimenCondominial: false,
      regimenDate: '',
      internalFinancingReady: false,
      updatedAt: new Date().toISOString()
    },
    units: []
  };
}

function normalizeDateString(raw) {
  const value = String(raw || '').trim();
  if (!value) return '';
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return '';
  return parsed.toISOString().slice(0, 10);
}

function daysUntil(dateText) {
  const normalized = normalizeDateString(dateText);
  if (!normalized) return null;
  const today = new Date();
  const target = new Date(`${normalized}T00:00:00`);
  const base = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return Math.floor((target - base) / (1000 * 60 * 60 * 24));
}

function evaluateUnitPriority(unit, project) {
  const paymentMethod = String(unit.paymentMethod || '').toLowerCase();
  const constructionStatus = String(unit.constructionStatus || '').toLowerCase();
  const handoverStatus = String(unit.handoverStatus || '').toLowerCase();
  const paymentStatus = String(unit.paymentStatus || '').toLowerCase();
  const hasFinanceIssues = Boolean(unit.hasFinanceIssues);
  const financeClearance = Boolean(unit.financeClearance);
  const appointmentConfirmed = Boolean(unit.appointmentConfirmed);
  const regimenReady = Boolean(project && project.regimenCondominial);
  const internalFinancingReady = Boolean(project && project.internalFinancingReady);
  const remainingDays = daysUntil(unit.deliveryDate);

  let score = 0;
  const blockers = [];

  if (paymentMethod === 'contado') score += 40;
  if (paymentMethod === 'credito_bancario') score += 30;
  if (paymentMethod === 'financiamiento_interno') {
    score -= 40;
    if (!internalFinancingReady) blockers.push('Financiamiento interno no disponible');
  }

  if (hasFinanceIssues || !financeClearance) {
    score -= 30;
    blockers.push('Cliente con incidencias en finanzas');
  }

  if (constructionStatus !== 'listo') {
    blockers.push('Obra no reporta unidad lista');
  } else if (remainingDays !== null && remainingDays <= 15) {
    score += 20;
  }

  if (!regimenReady) {
    blockers.push('Regimen condominial pendiente');
  }

  if (appointmentConfirmed) score += 10;
  if (handoverStatus === 'entregado') score -= 5;
  if (paymentStatus === 'pagado') score -= 10;

  const readyForDelivery = blockers.length === 0;
  return { score, blockers, readyForDelivery };
}

function readOwnerServicesData() {
  const raw = readJson(OWNER_SERVICES_PATH, ownerServicesDefaultData());
  const safe = ownerServicesDefaultData();
  safe.project = { ...safe.project, ...(raw.project || {}) };
  safe.project.regimenCondominial = Boolean(safe.project.regimenCondominial);
  safe.project.internalFinancingReady = Boolean(safe.project.internalFinancingReady);
  safe.project.regimenDate = normalizeDateString(safe.project.regimenDate);
  safe.units = Array.isArray(raw.units) ? raw.units : [];
  return safe;
}

function enrichOwnerUnit(unit, project) {
  const merged = {
    unitId: String(unit.unitId || '').trim(),
    tower: String(unit.tower || '').trim(),
    clientName: String(unit.clientName || '').trim(),
    deliveryDate: normalizeDateString(unit.deliveryDate),
    constructionStatus: String(unit.constructionStatus || 'pendiente').toLowerCase(),
    financeClearance: Boolean(unit.financeClearance),
    hasFinanceIssues: Boolean(unit.hasFinanceIssues),
    financeNotes: String(unit.financeNotes || '').trim(),
    paymentMethod: String(unit.paymentMethod || 'contado').toLowerCase(),
    appointmentConfirmed: Boolean(unit.appointmentConfirmed),
    handoverStatus: String(unit.handoverStatus || 'pendiente').toLowerCase(),
    paymentStatus: String(unit.paymentStatus || 'pendiente').toLowerCase(),
    ownerServicesPriority: Number.isFinite(Number(unit.ownerServicesPriority))
      ? Math.max(0, Math.min(100, Number(unit.ownerServicesPriority)))
      : 50,
    notes: String(unit.notes || '').trim(),
    updatedAt: String(unit.updatedAt || new Date().toISOString())
  };

  const evalData = evaluateUnitPriority(merged, project);
  return {
    ...merged,
    score: evalData.score,
    blockers: evalData.blockers,
    readyForDelivery: evalData.readyForDelivery
  };
}

function ownerServicesDashboardSnapshot() {
  const data = readOwnerServicesData();
  const enrichedUnits = data.units.map((u) => enrichOwnerUnit(u, data.project));
  const prioritized = [...enrichedUnits].sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return a.unitId.localeCompare(b.unitId);
  });
  const blocked = prioritized.filter((u) => !u.readyForDelivery);
  const ready = prioritized.filter((u) => u.readyForDelivery);
  const delivered = prioritized.filter((u) => u.handoverStatus === 'entregado');
  const paid = prioritized.filter((u) => u.paymentStatus === 'pagado');

  const summary = {
    totalUnits: prioritized.length,
    readyForDelivery: ready.length,
    blocked: blocked.length,
    delivered: delivered.length,
    paid: paid.length
  };

  return { project: data.project, summary, prioritized, blocked, units: enrichedUnits };
}

function saveOwnerServicesData(data) {
  writeJson(OWNER_SERVICES_PATH, data);
}

function sanitizeOwnerUnitPayload(raw) {
  return {
    unitId: String(raw.unitId || '').trim(),
    tower: String(raw.tower || '').trim(),
    clientName: String(raw.clientName || '').trim(),
    deliveryDate: normalizeDateString(raw.deliveryDate),
    constructionStatus: String(raw.constructionStatus || 'pendiente').toLowerCase(),
    financeClearance: Boolean(raw.financeClearance),
    hasFinanceIssues: Boolean(raw.hasFinanceIssues),
    financeNotes: String(raw.financeNotes || '').trim(),
    paymentMethod: String(raw.paymentMethod || 'contado').toLowerCase(),
    appointmentConfirmed: Boolean(raw.appointmentConfirmed),
    handoverStatus: String(raw.handoverStatus || 'pendiente').toLowerCase(),
    paymentStatus: String(raw.paymentStatus || 'pendiente').toLowerCase(),
    ownerServicesPriority: Number.isFinite(Number(raw.ownerServicesPriority))
      ? Math.max(0, Math.min(100, Number(raw.ownerServicesPriority)))
      : 50,
    notes: String(raw.notes || '').trim(),
    updatedAt: new Date().toISOString()
  };
}

function normalizeHeaderKey(raw) {
  return String(raw || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function parseBooleanValue(raw) {
  if (typeof raw === 'boolean') return raw;
  const value = String(raw || '').trim().toLowerCase();
  if (!value) return false;
  return ['si', 'sí', 'yes', 'true', '1', 'x'].includes(value);
}

function normalizePaymentMethod(raw) {
  const value = normalizeHeaderKey(raw);
  if (['contado', 'cash'].includes(value)) return 'contado';
  if (['credito_bancario', 'credito', 'creditobancario', 'bancario'].includes(value)) return 'credito_bancario';
  if (['credito_reta', 'reta', 'credito_reta_14'].includes(value)) return 'financiamiento_interno';
  if (['financiamiento_interno', 'interno', 'credito_interno', 'financiamiento'].includes(value)) return 'financiamiento_interno';
  return 'contado';
}

function normalizeConstructionStatus(raw) {
  const value = normalizeHeaderKey(raw);
  if (['listo', 'ready'].includes(value)) return 'listo';
  if (['entregado', 'delivered'].includes(value)) return 'entregado';
  return 'pendiente';
}

function normalizeHandoverStatus(raw) {
  const value = normalizeHeaderKey(raw);
  if (['programado', 'scheduled'].includes(value)) return 'programado';
  if (['entregado', 'delivered'].includes(value)) return 'entregado';
  return 'pendiente';
}

function normalizePaymentStatus(raw) {
  const value = normalizeHeaderKey(raw);
  if (['pagado', 'paid'].includes(value)) return 'pagado';
  if (['en_proceso', 'proceso', 'processing'].includes(value)) return 'en_proceso';
  return 'pendiente';
}

function excelDateToIso(raw) {
  if (raw === null || raw === undefined || raw === '') return '';
  if (typeof raw === 'number') {
    const parsed = XLSX.SSF.parse_date_code(raw);
    if (!parsed) return '';
    const mm = String(parsed.m).padStart(2, '0');
    const dd = String(parsed.d).padStart(2, '0');
    return `${parsed.y}-${mm}-${dd}`;
  }
  return normalizeDateString(raw);
}

const OWNER_EXCEL_COLUMN_MAP = {
  unitId: ['departamento', 'depto', 'unit_id', 'unitid', 'unidad', 'no'],
  tower: ['tower', 'torre'],
  clientName: ['client_name', 'cliente', 'nombre_cliente', 'alias_razon_social'],
  firstName: ['nombre_s', 'nombre'],
  lastName1: ['apellido_1', 'apellido_paterno'],
  lastName2: ['apellido_2', 'apellido_materno'],
  deliveryDate: ['delivery_date', 'fecha_entrega', 'fecha_estimada_entrega'],
  deliveryDateFromWork: ['fecha_entrega_obra'],
  deliveryDateClient: ['fecha_entrega_cliente'],
  constructionStatus: ['construction_status', 'estatus_obra', 'obra_estatus'],
  financeClearance: ['finance_clearance', 'finanzas_liberado', 'finanzas_ok'],
  hasFinanceIssues: ['has_finance_issues', 'incidencias_finanzas', 'finanzas_incidencias'],
  statusSos: ['status_sos'],
  financeNotes: ['finance_notes', 'notas_finanzas', 'comentarios_finanzas'],
  paymentMethod: ['payment_method', 'metodo_pago', 'tipo_pago'],
  appointmentConfirmed: ['appointment_confirmed', 'cita_confirmada', 'cita_entrega_confirmada'],
  handoverStatus: ['handover_status', 'estatus_entrega', 'estado_entrega'],
  paymentStatus: ['payment_status', 'estatus_pago', 'estado_pago', 'estatus'],
  notes: ['notes', 'notas_owner_services', 'notas']
};

function getRowValueByAliases(row, aliases) {
  for (const alias of aliases) {
    if (Object.prototype.hasOwnProperty.call(row, alias)) {
      return row[alias];
    }
  }
  return '';
}

function normalizeExcelRows(rows) {
  const normalizedRows = [];
  for (const rawRow of rows) {
    const row = {};
    Object.entries(rawRow || {}).forEach(([key, value]) => {
      row[normalizeHeaderKey(key)] = value;
    });
    const unitId = String(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.unitId) || '').trim();
    if (!unitId) continue;
    const aliasOrCompany = String(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.clientName) || '').trim();
    const firstName = String(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.firstName) || '').trim();
    const lastName1 = String(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.lastName1) || '').trim();
    const lastName2 = String(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.lastName2) || '').trim();
    const fullName = [firstName, lastName1, lastName2].filter(Boolean).join(' ').trim();
    const clientName = aliasOrCompany || fullName;

    const statusSos = String(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.statusSos) || '').trim();
    const statusSosNorm = normalizeHeaderKey(statusSos);
    const statusSosHasIssue = ['pendiente', 'bloqueado', 'observacion', 'incidencia', 'error']
      .some((word) => statusSosNorm.includes(word));

    const financeClearanceRaw = getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.financeClearance);
    const hasFinanceIssuesRaw = getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.hasFinanceIssues);
    const hasExplicitFinanceClearance = String(financeClearanceRaw || '').trim() !== '';
    const hasExplicitFinanceIssues = String(hasFinanceIssuesRaw || '').trim() !== '';
    const financeClearance = hasExplicitFinanceClearance
      ? parseBooleanValue(financeClearanceRaw)
      : !statusSosHasIssue;
    const hasFinanceIssues = hasExplicitFinanceIssues
      ? parseBooleanValue(hasFinanceIssuesRaw)
      : statusSosHasIssue;

    const deliveryDateWork = excelDateToIso(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.deliveryDateFromWork));
    const deliveryDate = excelDateToIso(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.deliveryDate)) || deliveryDateWork;
    const deliveryDateClient = excelDateToIso(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.deliveryDateClient));
    const appointmentConfirmedRaw = getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.appointmentConfirmed);
    const appointmentConfirmed = String(appointmentConfirmedRaw || '').trim()
      ? parseBooleanValue(appointmentConfirmedRaw)
      : Boolean(deliveryDateClient);

    const notesFromComments = String(getRowValueByAliases(row, ['comentarios']) || '').trim();
    const ticket = String(getRowValueByAliases(row, ['ticket']) || '').trim();
    const mergedNotes = [notesFromComments, ticket ? `Ticket: ${ticket}` : ''].filter(Boolean).join(' | ');

    let handoverStatus = normalizeHandoverStatus(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.handoverStatus));
    if (handoverStatus === 'pendiente' && deliveryDateClient) handoverStatus = 'programado';

    normalizedRows.push(sanitizeOwnerUnitPayload({
      unitId,
      tower: getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.tower),
      clientName,
      deliveryDate,
      constructionStatus: normalizeConstructionStatus(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.constructionStatus)),
      financeClearance,
      hasFinanceIssues,
      financeNotes: getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.financeNotes) || statusSos,
      paymentMethod: normalizePaymentMethod(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.paymentMethod)),
      appointmentConfirmed,
      handoverStatus,
      paymentStatus: normalizePaymentStatus(getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.paymentStatus)),
      notes: getRowValueByAliases(row, OWNER_EXCEL_COLUMN_MAP.notes) || mergedNotes
    }));
  }
  return normalizedRows;
}

function rowsFromSheetWithHeaderRow(sheet, headerRowNumber) {
  const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const headerIndex = Math.max(0, Number(headerRowNumber || 1) - 1);
  const headerRow = Array.isArray(grid[headerIndex]) ? grid[headerIndex] : [];
  const headers = headerRow.map((h) => String(h || '').trim());
  const rows = [];

  for (let i = headerIndex + 1; i < grid.length; i += 1) {
    const values = Array.isArray(grid[i]) ? grid[i] : [];
    const record = {};
    let hasAnyValue = false;
    for (let c = 0; c < headers.length; c += 1) {
      const key = headers[c];
      if (!key) continue;
      const value = values[c] === undefined ? '' : values[c];
      if (String(value).trim() !== '') hasAnyValue = true;
      record[key] = value;
    }
    if (hasAnyValue) rows.push(record);
  }

  return rows;
}

function persistSubmission(formatId, formatName, payload) {
  const current = readJson(SUBMISSIONS_PATH, []);
  const createdAt = new Date().toISOString();
  const record = {
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`,
    createdAt,
    formatId,
    formatName,
    payload
  };
  current.push(record);
  writeJson(SUBMISSIONS_PATH, current);
  // Keep latest payload for preview routes
  writeJson(DATA_PATH, payload);
  return record;
}

ensureDataFiles();

function renderTemplate(templateName, data, options = {}) {
  const filePath = path.join(TEMPLATE_DIR, templateName);
  let html = fs.readFileSync(filePath, 'utf-8');

  const watermarkUrl = options.baseUrl
    ? `${options.baseUrl}/assets/watermark.png`
    : '/assets/watermark.png';

  const controls = options.showControls
    ? `<button class="btn" type="button" onclick="history.back()">Regresar</button><button class="btn" onclick="window.location.href='${options.downloadUrl}'">Descargar PDF</button>`
    : '';

  html = html.replace(/{{watermarkUrl}}/g, watermarkUrl);
  html = html.replace(/{{controls}}/g, controls);
  html = html.replace(/{{bodyClass}}/g, options.bodyClass || 'interactive');

  // Simple {{key}} replacement using provided data
  Object.entries(data || {}).forEach(([key, value]) => {
    const re = new RegExp(`{{\\s*${key}\\s*}}`, 'g');
    html = html.replace(re, String(value));
  });

  return html;
}

app.get('/login', (req, res) => {
  if (req.isAuthenticated && req.isAuthenticated()) return res.redirect('/');
  const authReady = Boolean(GOOGLE_CLIENT_ID && GOOGLE_CLIENT_SECRET);
  const error = req.query && req.query.error === 'domain' ? `Solo cuentas @${ALLOWED_DOMAIN}.` : '';
  res.send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Login SIMCA</title>
  <style>
    body{font-family:Arial,sans-serif;background:#f4f1e8;margin:0;display:grid;place-items:center;min-height:100vh;padding:20px;}
    .card{width:min(540px,100%);background:#fff;border:1px solid #dcd7cb;border-radius:14px;padding:22px;}
    h1{margin:0 0 8px;font-size:26px;}
    p{margin:0 0 14px;color:#5f5f5f;}
    .btn{display:inline-block;padding:10px 14px;border:1px solid #bdb8a9;border-radius:10px;background:#fff;color:#111;text-decoration:none;font-weight:600;}
    .warn{margin-top:10px;color:#9d2a2a;font-size:14px;}
    .muted{margin-top:8px;font-size:12px;color:#6f6a5b;}
  </style></head><body>
    <div class="card">
      <h1>Acceso SIMCA</h1>
      <p>Inicia sesión con Google usando tu cuenta <strong>@${ALLOWED_DOMAIN}</strong>.</p>
      ${authReady ? '<a class="btn" href="/auth/google">Entrar con Google</a>' : '<p class="warn">Falta configurar GOOGLE_CLIENT_ID y GOOGLE_CLIENT_SECRET.</p>'}
      ${error ? `<p class="warn">${error}</p>` : ''}
      <p class="muted">Gerente ventas permitido: ${GERENTE_EMAIL}</p>
    </div>
  </body></html>`);
});

if (AUTH_READY) {
  app.get('/auth/google', passport.authenticate('google', {
    scope: ['profile', 'email'],
    hd: ALLOWED_DOMAIN
  }));

  app.get('/auth/google/callback',
    passport.authenticate('google', { failureRedirect: '/login?error=domain' }),
    (req, res) => res.redirect('/')
  );
} else {
  app.get('/auth/google', (req, res) => res.redirect('/login'));
  app.get('/auth/google/callback', (req, res) => res.redirect('/login'));
}

app.get('/logout', (req, res, next) => {
  req.logout((err) => {
    if (err) return next(err);
    req.session.destroy(() => {
      res.clearCookie('simca.sid');
      res.redirect('/login');
    });
  });
});

app.get('/', requireAuth, (req, res) => {
  const currentEmail = String(req.user && req.user.email || '').toLowerCase();
  const isGerente = currentEmail === GERENTE_EMAIL;
  const ownerServicesCard = isGerente ? `
        <a class="card" href="/owner-services">
          <span class="tag">Módulo</span>
          <h2 class="name">Owner Services</h2>
          <p class="desc">Prioriza entregas y coordina obra, jurídico y finanzas.</p>
        </a>` : '';
  const gerenteCard = isGerente ? `
        <a class="card" href="/gerente-ventas">
          <span class="tag">Módulo</span>
          <h2 class="name">Gerente Ventas</h2>
          <p class="desc">Acceso directo a herramientas de planos, edición y descarga PDF.</p>
        </a>` : '';
  const roiMasterPanel = isGerente ? `
      <section class="master-box master-box-mini">
        <h3>CSV Maestro</h3>
        <div class="master-row">
          <input id="roiMasterCsvFile" type="file" accept=".csv" />
          <button id="roiMasterCsvBtn" type="button">Actualizar</button>
        </div>
        <p id="roiMasterCsvStatus" class="master-status"></p>
      </section>` : '';
  res.send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Backend SIMCA</title>
  <style>
    :root{--bg:#f4f1e8;--card:#ffffff;--ink:#1a1a1a;--muted:#5f5f5f;--accent:#ffe816;--line:#d8d1c1;}
    *{box-sizing:border-box}
    body{font-family: Arial, sans-serif; margin:0; background:var(--bg); color:var(--ink);}
    .wrap{max-width:1000px; margin:0 auto; padding:32px 20px 60px;}
    h1{margin:0 0 6px; font-size:30px;}
    .sub{margin:0 0 26px; color:var(--muted);}
    .grid{display:grid; grid-template-columns:repeat(3,1fr); gap:14px;}
    .card{display:block; background:var(--card); border:1px solid #dcd7cb; border-radius:14px; padding:18px; text-decoration:none; color:inherit;}
    .card:hover{border-color:#b9b39f;}
    .card.maintenance{position:relative; overflow:hidden;}
    .card.maintenance .ribbon{
      position:absolute;
      top:14px;
      right:-70px;
      transform:rotate(34deg);
      width:260px;
      text-align:center;
      background:repeating-linear-gradient(-45deg,#ffe885,#ffe885 12px,#fff5bf 12px,#fff5bf 24px);
      border:1px solid #d6a700;
      color:#3e3300;
      font-size:11px;
      font-weight:800;
      letter-spacing:.08em;
      text-transform:uppercase;
      padding:5px 10px;
      z-index:2;
      pointer-events:none;
    }
    .tag{display:inline-block; font-size:12px; font-weight:700; background:var(--accent); padding:4px 8px; border-radius:999px; margin-bottom:10px;}
    .name{font-size:20px; margin:0 0 8px;}
    .desc{margin:0; color:var(--muted);}
    .top{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;}
    .left-head{display:flex;flex-direction:column;gap:10px;align-items:flex-start;}
    .user{font-size:13px;color:#5f5f5f;}
    .logout{display:inline-block;padding:8px 10px;border:1px solid #bdb8a9;border-radius:10px;background:#fff;color:#111;text-decoration:none;font-size:13px;font-weight:600;}
    .master-box{width:min(430px,100%);background:#fff;border:1px solid #d8d1c1;border-radius:12px;padding:10px 12px;}
    .master-box h3{margin:0 0 6px;font-size:14px;}
    .master-box-mini{width:230px;padding:6px 8px;border-radius:10px;}
    .master-box-mini h3{font-size:11px;margin:0 0 4px;}
    .master-box-mini .master-row{gap:6px;}
    .master-box-mini .master-row input{font-size:10px;}
    .master-box-mini .master-row button{padding:5px 8px;font-size:10px;border-radius:8px;}
    .master-box-mini .master-status{min-height:12px;margin:5px 0 0;font-size:10px;line-height:1.2;}
    .master-note{margin:0 0 8px;color:#5f5f5f;font-size:12px;}
    .master-row{display:flex;gap:8px;align-items:center;}
    .master-row input{flex:1;min-width:0;}
    .master-row button{padding:8px 12px;border:1px solid #111;border-radius:10px;background:#ffe816;color:#111;font-weight:700;cursor:pointer;}
    .master-status{min-height:16px;margin:8px 0 0;font-size:12px;color:#5f5f5f;}
    @media (max-width:900px){
      .wrap{padding:20px 14px 36px;}
      .top{flex-direction:column;align-items:flex-start;}
      .grid{grid-template-columns:1fr;}
    }
    @media (max-width:640px){
      h1{font-size:24px;}
      .sub{margin-bottom:16px;}
      .card{padding:14px;}
      .card.maintenance .ribbon{
        top:10px;
        right:10px;
        transform:none;
        width:auto;
        max-width:72%;
        font-size:10px;
        letter-spacing:.04em;
        padding:4px 8px;
      }
    }
  </style></head><body>
    <div class="wrap">
      <div class="top">
        <div class="left-head">
          <h1>Backend SIMCA</h1>
          <p class="sub">Panel principal de módulos.</p>
        </div>
        <div style="text-align:right;">
          <div class="user">${String(req.user && req.user.email || '')}</div>
          <a class="logout" href="/logout">Cerrar sesión</a>
          <div style="margin-top:8px;display:flex;justify-content:flex-end;">
            ${roiMasterPanel}
          </div>
        </div>
      </div>
      <div class="grid">
        <a class="card" href="/generador-faes">
          <span class="tag">Módulo</span>
          <h2 class="name">Generador FAES</h2>
          <p class="desc">Gestión y generación de documentos FAES.</p>
        </a>
        <a class="card maintenance" href="/plds">
          <span class="ribbon">Under Construction</span>
          <span class="tag">Módulo</span>
          <h2 class="name">PLDS</h2>
          <p class="desc">En mantenimiento. No usar temporalmente.</p>
        </a>
        <a class="card" href="/generador-roi">
          <span class="tag">Módulo</span>
          <h2 class="name">Generador ROI</h2>
          <p class="desc">Cálculo de retorno de inversión por unidad.</p>
        </a>
        ${ownerServicesCard}
        ${gerenteCard}
      </div>
    </div>
    <script>
      (function () {
        const input = document.getElementById('roiMasterCsvFile');
        const btn = document.getElementById('roiMasterCsvBtn');
        const status = document.getElementById('roiMasterCsvStatus');
        if (!input || !btn || !status) return;

        btn.addEventListener('click', async function () {
          const file = input.files && input.files[0];
          if (!file) {
            status.textContent = 'Selecciona un archivo CSV.';
            return;
          }
          if (!String(file.name || '').toLowerCase().endsWith('.csv')) {
            status.textContent = 'El archivo debe ser .csv';
            return;
          }
          btn.disabled = true;
          status.textContent = 'Subiendo CSV maestro...';
          try {
            const base64 = await new Promise(function (resolve, reject) {
              const r = new FileReader();
              r.onload = function () { resolve(r.result); };
              r.onerror = reject;
              r.readAsDataURL(file);
            });
            const res = await fetch('/api/roi/master-csv', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              credentials: 'same-origin',
              body: JSON.stringify({ fileName: file.name, base64: base64 })
            });
            const data = await res.json().catch(function () { return {}; });
            if (!res.ok) {
              status.textContent = data.error || 'No se pudo actualizar el CSV maestro.';
              return;
            }
            status.textContent = 'CSV maestro actualizado correctamente.';
          } catch (err) {
            status.textContent = 'Error al subir el CSV maestro.';
          } finally {
            btn.disabled = false;
          }
        });
      })();
    </script>
  </body></html>`);
});

app.use('/legacy', requireAuth);
app.use('/generador-faes', requireAuth);
app.use('/plds', requireAuth);
app.use('/generador-roi', requireAuth);
app.use('/form', requireAuth);
app.use('/format', requireAuth);
app.use('/submissions', requireAuth);
app.use('/api/plds', requireAuth);
app.use('/api/roi', requireAuth);
app.use('/owner-services', requireGerente);
app.use('/api/owner-services', requireGerente);

app.use('/gerente-ventas', requireGerente);
app.use('/plano-interactivo', requireGerente);
app.use('/plano-ventas', requireGerente);
app.use('/plano-descargar', requireGerente);
app.use('/api/plano-ventas', requireGerente);

app.get('/legacy/formatos', (req, res) => {
  const items = Object.entries(formats)
    .map(([id, f]) => `<li><a href="/format/${id}">${f.name}</a></li>`)
    .join('');

  res.send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <title>Formatos FR-VEN</title>
  <style>
    body{font-family: Arial, sans-serif; padding:24px;}
    li{margin:8px 0;}
  </style></head><body>
    <h1>Formatos disponibles</h1>
    <ul>${items}</ul>
    <p><a href="/">Volver al panel</a></p>
  </body></html>`);
});

app.get('/generador-faes', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'faes', 'index.html'));
});

app.get('/plds', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'plds', 'index.html'));
});

function renderPldsPlaceholder(title, res) {
  res.send(`<!doctype html><html lang="es"><head><meta charset="utf-8"><title>${title}</title></head><body style="font-family:Arial;padding:24px;background:#e0ded6"><button type="button" onclick="history.back()" style="margin-bottom:12px;padding:8px 12px;border:1px solid #333;border-radius:8px;background:#fff;cursor:pointer;">Regresar</button><h1>${title}</h1><p>Sección en construcción.</p><p><a href="/plds">Volver a PLDS</a></p></body></html>`);
}

app.get('/plds/cliente-fideicomiso-bilingue', (req, res) => {
  renderPldsPlaceholder('Cliente Fideicomiso Bilingüe', res);
});

app.get('/plds/cliente-fideicomiso-espanol', (req, res) => {
  renderPldsPlaceholder('Cliente Fideicomiso Español', res);
});

app.get('/plds/cliente-fideicomiso-empresarial-bilingue', (req, res) => {
  renderPldsPlaceholder('Cliente Fideicomiso Empresarial Bilingüe', res);
});

app.get('/plds/cliente-fideicomiso-empresarial-espanol', (req, res) => {
  renderPldsPlaceholder('Cliente Fideicomiso Empresarial Español', res);
});

app.get('/plds/cliente-extranjera-persona-moral', (req, res) => {
  renderPldsPlaceholder('Cliente Extranjera Persona Moral', res);
});

app.get('/plds/cliente-nacional-persona-fisica', (req, res) => {
  renderPldsPlaceholder('Cliente Nacional Persona Física', res);
});

app.get('/plds/cliente-nacional-persona-moral', (req, res) => {
  renderPldsPlaceholder('Cliente Nacional Persona Moral', res);
});

app.get('/generador-roi', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'generador-roi.html'));
});

app.get('/api/roi/master-csv', (req, res) => {
  if (!fs.existsSync(ROI_MASTER_CSV_PATH)) {
    return res.status(404).json({ error: 'CSV maestro no configurado' });
  }
  res.type('text/csv; charset=utf-8');
  return res.send(fs.readFileSync(ROI_MASTER_CSV_PATH, 'utf-8'));
});

app.post('/api/roi/master-csv', requireGerente, (req, res) => {
  try {
    const fileName = String(req.body?.fileName || '').toLowerCase();
    const base64Content = String(req.body?.base64 || '');
    if (!fileName.endsWith('.csv') || !base64Content) {
      return res.status(400).json({ error: 'Archivo CSV inválido' });
    }
    const commaIndex = base64Content.indexOf(',');
    const payload = commaIndex >= 0 ? base64Content.slice(commaIndex + 1) : base64Content;
    const buffer = Buffer.from(payload, 'base64');
    if (!buffer.length) {
      return res.status(400).json({ error: 'Contenido CSV vacío' });
    }
    fs.writeFileSync(ROI_MASTER_CSV_PATH, buffer);
    return res.json({ ok: true, filePath: ROI_MASTER_CSV_PATH });
  } catch (err) {
    return res.status(500).json({
      error: 'No se pudo guardar el CSV maestro',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/owner-services', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'owner-services.html'));
});

app.get('/plano-interactivo', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'plano-interactivo.html'));
});

app.get('/plano-ventas', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'plano-ventas.html'));
});

app.get('/gerente-ventas', (req, res) => {
  res.send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Gerente Ventas</title>
  <style>
    :root{--bg:#f4f1e8;--card:#ffffff;--ink:#1a1a1a;--muted:#5f5f5f;--accent:#ffe816;}
    *{box-sizing:border-box}
    body{font-family:Arial,sans-serif;margin:0;background:var(--bg);color:var(--ink);}
    .wrap{max-width:1040px;margin:0 auto;padding:28px 20px 52px;}
    h1{margin:0 0 8px;font-size:30px;}
    .sub{margin:0 0 24px;color:var(--muted);}
    .grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;}
    .card{display:block;background:var(--card);border:1px solid #dcd7cb;border-radius:14px;padding:18px;text-decoration:none;color:inherit;}
    .card:hover{border-color:#b9b39f;}
    .tag{display:inline-block;font-size:12px;font-weight:700;background:var(--accent);padding:4px 8px;border-radius:999px;margin-bottom:10px;}
    .name{font-size:21px;margin:0 0 8px;}
    .desc{margin:0;color:var(--muted);}
    .back{
      display:inline-block;
      text-decoration:none;
      color:#111;
      border:1px solid var(--line);
      background:#fff;
      border-radius:10px;
      padding:10px 12px;
      font-size:13px;
      font-weight:600;
      white-space:nowrap;
      cursor:pointer;
    }
    .back:hover{border-color:#b6ad98;}
    .hero{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;margin-bottom:20px;}
    @media (max-width:900px){.grid{grid-template-columns:1fr;}.hero{align-items:flex-start;flex-direction:column;}}
  </style></head><body>
    <div class="wrap">
      <div class="hero">
        <div>
          <h1>Gerente Ventas</h1>
          <p class="sub">Módulo central para flujo de pisos, edición visual y descarga PDF.</p>
        </div>
        <a class="back" href="/">Regresar</a>
      </div>
      <div class="grid">
        ${DEVELOPMENTS.map((dev, idx) => `
        <a class="card" href="/gerente-ventas/${dev.slug}">
          <span class="tag">Desarrollo ${idx + 1}</span>
          <h2 class="name">${dev.name}</h2>
          <p class="desc">Abrir herramientas de plano para ${dev.name}.</p>
        </a>`).join('')}
      </div>
    </div>
  </body></html>`);
});

app.get('/gerente-ventas/:devSlug', (req, res) => {
  const dev = getRequestedDevelopment(req);
  res.send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Gerente Ventas - ${dev.name}</title>
  <style>
    :root{--bg:#f4f1e8;--card:#ffffff;--ink:#1a1a1a;--muted:#5f5f5f;--accent:#ffe816;--line:#d8d1c1;}
    *{box-sizing:border-box}
    body{font-family:Arial,sans-serif;margin:0;background:var(--bg);color:var(--ink);}
    .wrap{max-width:1040px;margin:0 auto;padding:28px 20px 52px;}
    h1{margin:0 0 8px;font-size:30px;}
    .sub{margin:0 0 24px;color:var(--muted);}
    .grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;}
    .card{display:block;background:var(--card);border:1px solid #dcd7cb;border-radius:14px;padding:18px;text-decoration:none;color:inherit;}
    .card:hover{border-color:#b9b39f;}
    .tag{display:inline-block;font-size:12px;font-weight:700;background:var(--accent);padding:4px 8px;border-radius:999px;margin-bottom:10px;}
    .name{font-size:21px;margin:0 0 8px;}
    .desc{margin:0;color:var(--muted);}
    .back{
      display:inline-block;text-decoration:none;color:#111;border:1px solid var(--line);background:#fff;border-radius:10px;
      padding:10px 12px;font-size:13px;font-weight:600;white-space:nowrap;cursor:pointer;
    }
    .back:hover{border-color:#b6ad98;}
    .hero{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;margin-bottom:20px;}
    @media (max-width:900px){.grid{grid-template-columns:1fr;}.hero{align-items:flex-start;flex-direction:column;}}
  </style></head><body>
    <div class="wrap">
      <div class="hero">
        <div>
          <h1>${dev.name}</h1>
          <p class="sub">Selecciona herramienta del desarrollo.</p>
        </div>
        <a class="back" href="/gerente-ventas">Regresar</a>
      </div>
      <div class="grid">
        <a class="card" href="/gerente-ventas/${dev.slug}/plano-interactivo?dev=${dev.slug}">
          <span class="tag">A</span>
          <h2 class="name">Plano Interactivo</h2>
          <p class="desc">Marcado de unidades por piso y guardado base.</p>
        </a>
        <a class="card" href="/gerente-ventas/${dev.slug}/plano-ventas?dev=${dev.slug}">
          <span class="tag">B</span>
          <h2 class="name">Plano Ventas (Editor)</h2>
          <p class="desc">Editor visual completo de PDF, tabla y estado de unidades.</p>
        </a>
        <a class="card" href="/gerente-ventas/${dev.slug}/plano-descargar?dev=${dev.slug}">
          <span class="tag">C</span>
          <h2 class="name">Plano Descargar</h2>
          <p class="desc">Descarga rápida del PDF con carga automática de datos.</p>
        </a>
      </div>
    </div>
  </body></html>`);
});

app.get('/gerente-ventas/:devSlug/plano-interactivo', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'plano-interactivo.html'));
});

app.get('/gerente-ventas/:devSlug/plano-ventas', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'plano-ventas.html'));
});

app.get('/gerente-ventas/:devSlug/plano-descargar', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'plano-descargar.html'));
});

app.get('/plano-descargar', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'plano-descargar.html'));
});

app.get('/api/plano-ventas/default-excel', (req, res) => {
  const dev = getRequestedDevelopment(req);
  const devDir = path.join(DEVELOPMENTS_DIR, dev.slug);
  const devFiles = [];
  try {
    const names = fs.readdirSync(devDir);
    names.forEach((name) => {
      if (/\.(xls|xlsx|csv)$/i.test(name)) {
        const fullPath = path.join(devDir, name);
        let mtimeMs = 0;
        try { mtimeMs = fs.statSync(fullPath).mtimeMs || 0; } catch {}
        devFiles.push({ fullPath, name, mtimeMs });
      }
    });
  } catch {}
  const canonicalNames = new Set([
    'inventariomaestrowix.xls',
    'inventariomaestrowix.xlsx',
    'inventariomaestrowix.csv'
  ]);
  const canonicalDevFiles = devFiles
    .filter((f) => canonicalNames.has(String(f.name || '').toLowerCase()))
    .sort((a, b) => b.mtimeMs - a.mtimeMs);
  const latestDevFiles = devFiles
    .filter((f) => !canonicalNames.has(String(f.name || '').toLowerCase()))
    .sort((a, b) => b.mtimeMs - a.mtimeMs);
  const candidates = [
    ...canonicalDevFiles.map((f) => f.fullPath),
    ...latestDevFiles.map((f) => f.fullPath),
    // Shared global master CSV uploaded from Home (CSV Maestro).
    ROI_MASTER_CSV_PATH,
    path.join(DEVELOPMENTS_DIR, dev.slug, 'INVENTARIOMAESTROWIX.xls'),
    path.join(DEVELOPMENTS_DIR, dev.slug, 'INVENTARIOMAESTROWIX.xlsx'),
    path.join(DEVELOPMENTS_DIR, dev.slug, 'INVENTARIOMAESTROWIX.csv'),
    path.join(os.homedir(), 'Downloads', 'INVENTARIOMAESTROWIX.xls'),
    path.join(os.homedir(), 'Downloads', 'INVENTARIOMAESTROWIX.xlsx'),
    path.join(os.homedir(), 'Downloads', 'INVENTARIOMAESTROWIX.csv')
  ];
  const dedupedCandidates = [...new Set(candidates)];
  const filePath = dedupedCandidates.find((candidate) => fs.existsSync(candidate));
  if (!filePath) {
    return res.status(404).json({
      error: 'Archivo no encontrado',
      expectedPath: path.join(DEVELOPMENTS_DIR, dev.slug, 'INVENTARIOMAESTROWIX.xlsx'),
      dev: dev.slug
    });
  }
  res.sendFile(filePath);
});

app.post('/api/plano-ventas/save-excel', (req, res) => {
  try {
    const dev = getRequestedDevelopment(req);
    const rawFileName = sanitizeExcelFileName(req.body?.fileName);
    const base64Content = String(req.body?.base64 || '');
    if (!rawFileName || !base64Content) {
      return res.status(400).json({ error: 'Archivo Excel inválido' });
    }
    const commaIndex = base64Content.indexOf(',');
    const payload = commaIndex >= 0 ? base64Content.slice(commaIndex + 1) : base64Content;
    const buffer = Buffer.from(payload, 'base64');
    if (!buffer.length) {
      return res.status(400).json({ error: 'Contenido vacío' });
    }
    const devDir = path.join(DEVELOPMENTS_DIR, dev.slug);
    if (!fs.existsSync(devDir)) fs.mkdirSync(devDir, { recursive: true });

    const ext = (rawFileName.split('.').pop() || '').toLowerCase();
    const canonicalName = `INVENTARIOMAESTROWIX.${ext}`;
    const canonicalPath = path.join(devDir, canonicalName);
    fs.writeFileSync(canonicalPath, buffer);

    const originalPath = path.join(devDir, rawFileName);
    if (originalPath !== canonicalPath) {
      fs.writeFileSync(originalPath, buffer);
    }

    return res.json({
      ok: true,
      dev: dev.slug,
      fileName: canonicalName,
      filePath: canonicalPath
    });
  } catch (err) {
    return res.status(500).json({
      error: 'No se pudo guardar Excel en carpeta del desarrollo',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/api/plano-ventas/default-json', (req, res) => {
  const dev = getRequestedDevelopment(req);
  const downloadsDir = path.join(os.homedir(), 'Downloads');
  const floorDirs = getDevelopmentFloorSearchDirs(dev.slug);
  let filePath = '';

  const requestName = typeof req.query.name === 'string' ? req.query.name.trim() : '';
  const safeName = requestName && !requestName.includes('/') && !requestName.includes('\\') ? requestName : '';
  if (safeName) {
    floorDirs.some((dir) => {
      const candidate = path.join(dir, safeName);
      if (fs.existsSync(candidate)) {
        filePath = candidate;
        return true;
      }
      return false;
    });
    if (!filePath) {
      const candidate = path.join(downloadsDir, safeName);
      if (fs.existsSync(candidate)) filePath = candidate;
    }
  } else {
    floorDirs.some((dir) => {
      const upperCandidate = path.join(dir, 'unidades-marcadas.JSON');
      const lowerCandidate = path.join(dir, 'unidades-marcadas.json');
      if (fs.existsSync(upperCandidate)) {
        filePath = upperCandidate;
        return true;
      }
      if (fs.existsSync(lowerCandidate)) {
        filePath = lowerCandidate;
        return true;
      }
      return false;
    });
  }

  if (!filePath && safeName) {
    floorDirs.some((dir) => {
      const files = listFloorJsonFiles(dir);
      const preferred = files.find((name) => name.toLowerCase() === safeName.toLowerCase());
      if (preferred) {
        filePath = path.join(dir, preferred);
        return true;
      }
      return false;
    });
  }

  if (!filePath && !safeName) {
    floorDirs.some((dir) => {
      const files = listFloorJsonFiles(dir);
      if (files[0]) {
        filePath = path.join(dir, files[0]);
        return true;
      }
      return false;
    });
  }

  if (!filePath && safeName) {
    try {
      const files = fs.readdirSync(downloadsDir);
      const preferred = files.find((name) => name.toLowerCase() === safeName.toLowerCase());
      if (preferred) filePath = path.join(downloadsDir, preferred);
    } catch {}
  }

  if (!filePath && !safeName) {
    try {
      const files = fs.readdirSync(downloadsDir)
        .filter((name) => FLOOR_JSON_FILE_RE.test(name))
        .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
      if (files[0]) filePath = path.join(downloadsDir, files[0]);
    } catch {}
  }

  if (!filePath || !fs.existsSync(filePath)) {
    return res.status(404).json({
      error: 'Archivo no encontrado',
      expectedPath: safeName
        ? path.join(getDevelopmentFloorDir(dev.slug), safeName)
        : path.join(getDevelopmentFloorDir(dev.slug), 'unidades-marcadas.json'),
      dev: dev.slug
    });
  }

  res.sendFile(filePath);
});

app.get('/api/plano-ventas/default-json-files', (req, res) => {
  const dev = getRequestedDevelopment(req);
  const downloadsDir = path.join(os.homedir(), 'Downloads');
  const floorDirs = getDevelopmentFloorSearchDirs(dev.slug);
  try {
    for (const dir of floorDirs) {
      const files = listFloorJsonFiles(dir);
      if (files.length) {
        return res.json({ files, sourceDir: dir, dev: dev.slug });
      }
    }

    const files = fs.readdirSync(downloadsDir)
      .filter((name) => FLOOR_JSON_FILE_RE.test(name))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
    res.json({ files, sourceDir: downloadsDir, dev: dev.slug });
  } catch (err) {
    res.status(500).json({ error: 'No se pudo leer carpeta de JSONs', details: err.message });
  }
});

app.get('/api/plano-ventas/version-json-files', (req, res) => {
  const dev = getRequestedDevelopment(req);
  const floorDirs = getDevelopmentFloorSearchDirs(dev.slug);
  try {
    const entries = [];
    floorDirs.forEach((dir) => {
      try {
        fs.readdirSync(dir)
          .filter((name) => {
            const lower = String(name || '').toLowerCase();
            if (!lower.endsWith('.json')) return false;
            if (/^version-.*\.json$/i.test(name)) return true;
            return FLOOR_JSON_FILE_RE.test(name);
          })
          .forEach((name) => {
            const fullPath = path.join(dir, name);
            let mtimeMs = 0;
            try {
              mtimeMs = fs.statSync(fullPath).mtimeMs || 0;
            } catch {}
            entries.push({ name, dir, mtimeMs });
          });
      } catch {}
    });

    const deduped = [];
    const seen = new Set();
    entries
      .sort((a, b) => b.mtimeMs - a.mtimeMs || a.name.localeCompare(b.name))
      .forEach((entry) => {
        const key = String(entry.name || '').toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        deduped.push(entry);
      });
    res.json({
      files: deduped.map((e) => e.name),
      sourceDir: floorDirs.join(', '),
      dev: dev.slug
    });
  } catch (err) {
    res.status(500).json({ error: 'No se pudo listar versiones JSON', details: err.message });
  }
});

app.get('/api/plano-ventas/json-files', (req, res) => {
  const dev = getRequestedDevelopment(req);
  const floorDirs = getDevelopmentFloorSearchDirs(dev.slug);
  try {
    const entries = [];
    floorDirs.forEach((dir) => {
      try {
        fs.readdirSync(dir)
          .filter((name) => String(name || '').toLowerCase().endsWith('.json'))
          .forEach((name) => {
            const fullPath = path.join(dir, name);
            let stat = null;
            try {
              stat = fs.statSync(fullPath);
            } catch {}
            if (!stat || !stat.isFile()) return;
            entries.push({
              name,
              dir,
              mtimeMs: stat.mtimeMs || 0
            });
          });
      } catch {}
    });
    const deduped = [];
    const seen = new Set();
    entries
      .sort((a, b) => b.mtimeMs - a.mtimeMs || a.name.localeCompare(b.name))
      .forEach((entry) => {
        const key = String(entry.name || '').toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        deduped.push(entry);
      });
    res.json({
      files: deduped.map((e) => e.name),
      sourceDir: floorDirs.join(', '),
      dev: dev.slug
    });
  } catch (err) {
    res.status(500).json({ error: 'No se pudo listar JSONs del desarrollo', details: err.message });
  }
});

app.delete('/api/plano-ventas/json-files', (req, res) => {
  const dev = getRequestedDevelopment(req);
  const nameRaw = typeof req.query.name === 'string' ? req.query.name.trim() : '';
  const safeName = nameRaw && !nameRaw.includes('/') && !nameRaw.includes('\\')
    ? path.basename(nameRaw)
    : '';
  if (!safeName || !safeName.toLowerCase().endsWith('.json')) {
    return res.status(400).json({ error: 'Nombre de archivo JSON inválido' });
  }

  const floorDirs = getDevelopmentFloorSearchDirs(dev.slug);
  const deletedPaths = [];
  floorDirs.forEach((dir) => {
    const target = path.join(dir, safeName);
    try {
      if (fs.existsSync(target) && fs.statSync(target).isFile()) {
        fs.unlinkSync(target);
        deletedPaths.push(target);
      }
    } catch {}
  });

  if (!deletedPaths.length) {
    return res.status(404).json({
      error: 'Archivo no encontrado para borrar',
      dev: dev.slug,
      fileName: safeName
    });
  }

  return res.json({
    ok: true,
    dev: dev.slug,
    fileName: safeName,
    deletedPaths
  });
});

app.get('/api/plano-ventas/default-json-merged', (req, res) => {
  const dev = getRequestedDevelopment(req);
  const downloadsDir = path.join(os.homedir(), 'Downloads');
  const floorDirs = getDevelopmentFloorSearchDirs(dev.slug);
  try {
    let sourceDirs = floorDirs;
    let filesByName = new Map();

    for (const dir of floorDirs) {
      const files = listFloorJsonFiles(dir);
      files.forEach((name) => {
        const key = name.toLowerCase();
        if (!filesByName.has(key)) {
          filesByName.set(key, { name, dir });
        }
      });
    }

    if (!filesByName.size) {
      sourceDirs = [downloadsDir];
      const downloadFiles = listFloorJsonFiles(downloadsDir);
      downloadFiles.forEach((name) => {
        filesByName.set(name.toLowerCase(), { name, dir: downloadsDir });
      });
    }

    const files = Array.from(filesByName.values())
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));

    const floors = [];
    const loadedFiles = [];
    let showUnitLabels = true;
    files.forEach(({ name, dir }) => {
      const filePath = path.join(dir, name);
      try {
        const raw = fs.readFileSync(filePath, 'utf-8');
        const parsed = JSON.parse(raw);
        if (parsed && parsed.showUnitLabels === false) showUnitLabels = false;
        const payloadFloors = Array.isArray(parsed)
          ? parsed
          : (parsed && Array.isArray(parsed.floors) ? parsed.floors : (parsed && parsed.imageDataUrl ? [parsed] : []));
        if (payloadFloors.length) {
          floors.push(...payloadFloors);
          loadedFiles.push(path.relative(DATA_DIR, filePath).replace(/\\/g, '/'));
        }
      } catch {}
    });

    res.json({
      floors,
      loadedFiles,
      sourceDir: sourceDirs.join(', '),
      dev: dev.slug,
      showUnitLabels
    });
  } catch (err) {
    res.status(500).json({ error: 'No se pudo leer JSONs de carpeta control', details: err.message });
  }
});

app.post('/api/plano-ventas/save-json', (req, res) => {
  try {
    const dev = getRequestedDevelopment(req);
    const body = req.body || {};
    const payload = body.payload;
    const fileName = sanitizeJsonFileName(body.fileName);
    if (!payload || typeof payload !== 'object') {
      return res.status(400).json({ error: 'Payload JSON inválido' });
    }
    const floorDir = getDevelopmentFloorDir(dev.slug);
    if (!fs.existsSync(floorDir)) fs.mkdirSync(floorDir, { recursive: true });
    const filePath = path.join(floorDir, fileName);
    fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), 'utf-8');
    return res.json({
      ok: true,
      dev: dev.slug,
      fileName,
      filePath
    });
  } catch (err) {
    return res.status(500).json({
      error: 'No se pudo guardar JSON en carpeta del desarrollo',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.post('/api/plano-ventas/render-pdf', async (req, res) => {
  const html = typeof req.body?.html === 'string' ? req.body.html : '';
  const fileNamePrefix = String(req.body?.fileNamePrefix || 'CEIBA').trim().toUpperCase().replace(/[^A-Z0-9 ]+/g, '').slice(0, 40) || 'CEIBA';
  if (!html || html.length < 20) {
    return res.status(400).json({ error: 'HTML inválido para generar PDF.' });
  }

  let browser = null;
  try {
    try {
      browser = await launchPdfBrowser();
    } catch (launchErr) {
      const installed = installChromeIfMissing(launchErr);
      if (!installed) throw launchErr;
      browser = await launchPdfBrowser();
    }
    const page = await browser.newPage();
    page.setDefaultTimeout(120000);
    page.setDefaultNavigationTimeout(120000);
    await page.setContent(html, { waitUntil: 'domcontentloaded', timeout: 120000 });
    await page.emulateMediaType('print');
    try {
      await page.waitForFunction(() => window.__pdfReady === true, { timeout: 15000 });
    } catch {
      // Fallback: if the marker is missing, continue and generate PDF.
    }

    const pdfBuffer = await page.pdf({
      printBackground: true,
      preferCSSPageSize: true
    });

    const utfName = `${fileNamePrefix} INVENTARIO ／ INVENTORY.pdf`;
    const fallbackName = `${fileNamePrefix} INVENTARIO - INVENTORY.pdf`;
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${fallbackName}"; filename*=UTF-8''${encodeURIComponent(utfName)}`
    );
    return res.send(pdfBuffer);
  } catch (err) {
    log(`Error en /api/plano-ventas/render-pdf: ${err && err.stack ? err.stack : err}`);
    return res.status(500).json({
      error: 'No se pudo generar el PDF en el servidor.',
      details: err && err.message ? err.message : 'error desconocido'
    });
  } finally {
    if (browser) {
      try { await browser.close(); } catch {}
    }
  }
});

app.get('/form', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'form.html'));
});

app.get('/health', (req, res) => {
  res.status(200).json({ ok: true, ts: new Date().toISOString() });
});

app.get('/submissions', (req, res) => {
  const limitRaw = Number(req.query.limit || 50);
  const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(limitRaw, 500)) : 50;
  const all = readJson(SUBMISSIONS_PATH, []);
  const items = all.slice(-limit).reverse();
  res.json({ total: all.length, items });
});

// --- API modulo PLDS ---
app.get('/api/plds/health', (req, res) => {
  res.json({ ok: true, module: 'plds', ts: new Date().toISOString() });
});

app.get('/api/plds/formats', (req, res) => {
  const items = Object.entries(formats).map(([id, f]) => ({
    id,
    name: f.name,
    template: f.file,
    htmlRoute: `/format/${id}`,
    pdfGetRoute: `/format/${id}/pdf`,
    pdfPostRoute: `/format/${id}/pdf`
  }));
  res.json({ items });
});

app.get('/api/plds/submissions', (req, res) => {
  const limitRaw = Number(req.query.limit || 50);
  const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(limitRaw, 500)) : 50;
  const all = readJson(SUBMISSIONS_PATH, []);
  const items = all.slice(-limit).reverse();
  res.json({ total: all.length, items });
});

app.get('/api/plds/submissions/:id', (req, res) => {
  const all = readJson(SUBMISSIONS_PATH, []);
  const item = all.find((x) => x.id === req.params.id);
  if (!item) return res.status(404).json({ error: 'Submission no encontrado' });
  res.json(item);
});

// --- API modulo Owner Services ---
app.get('/api/owner-services/health', (req, res) => {
  res.json({ ok: true, module: 'owner-services', ts: new Date().toISOString() });
});

app.get('/api/owner-services/dashboard', (req, res) => {
  res.json(ownerServicesDashboardSnapshot());
});

app.put('/api/owner-services/project', (req, res) => {
  const current = readOwnerServicesData();
  const body = req.body || {};
  const updatedProject = {
    ...current.project,
    name: String(body.name || current.project.name || '').trim() || 'Proyecto principal',
    regimenCondominial: Boolean(body.regimenCondominial),
    regimenDate: normalizeDateString(body.regimenDate),
    internalFinancingReady: Boolean(body.internalFinancingReady),
    updatedAt: new Date().toISOString()
  };
  const next = { ...current, project: updatedProject };
  saveOwnerServicesData(next);
  res.json({ ok: true, project: next.project });
});

app.post('/api/owner-services/units', (req, res) => {
  const body = req.body || {};
  const incoming = sanitizeOwnerUnitPayload(body);
  const unitId = incoming.unitId;
  if (!unitId) {
    return res.status(400).json({ error: 'unitId es obligatorio' });
  }

  const current = readOwnerServicesData();
  const idx = current.units.findIndex((x) => String(x.unitId || '').trim() === unitId);
  if (idx >= 0) {
    current.units[idx] = { ...current.units[idx], ...incoming };
  } else {
    current.units.push(incoming);
  }

  saveOwnerServicesData(current);
  const snapshot = ownerServicesDashboardSnapshot();
  const saved = snapshot.units.find((x) => x.unitId === unitId);
  return res.json({ ok: true, unit: saved });
});

app.get('/api/owner-services/template-columns', (req, res) => {
  const columns = [
    'unit_id',
    'torre',
    'cliente',
    'fecha_entrega',
    'estatus_obra',
    'finanzas_liberado',
    'incidencias_finanzas',
    'notas_finanzas',
    'metodo_pago',
    'cita_confirmada',
    'estatus_entrega',
    'estatus_pago',
    'notas_owner_services'
  ];
  res.json({
    columns,
    hints: {
      estatus_obra: ['pendiente', 'listo', 'entregado'],
      metodo_pago: ['contado', 'credito_bancario', 'financiamiento_interno'],
      estatus_entrega: ['pendiente', 'programado', 'entregado'],
      estatus_pago: ['pendiente', 'en_proceso', 'pagado'],
      boolean_values: ['si', 'no']
    }
  });
});

app.get('/api/owner-services/template-csv', (req, res) => {
  const header = [
    'unit_id',
    'torre',
    'cliente',
    'fecha_entrega',
    'estatus_obra',
    'finanzas_liberado',
    'incidencias_finanzas',
    'notas_finanzas',
    'metodo_pago',
    'cita_confirmada',
    'estatus_entrega',
    'estatus_pago',
    'notas_owner_services'
  ].join(',');
  const example = [
    'T1-502',
    'Torre 1',
    'Juan Perez',
    '2026-02-20',
    'listo',
    'si',
    'no',
    'Sin incidencias',
    'contado',
    'si',
    'programado',
    'pendiente',
    'Cliente listo para firma'
  ].join(',');

  const csv = `${header}\n${example}\n`;
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename=owner-services-template.csv');
  res.send(csv);
});

app.post('/api/owner-services/import-excel', (req, res) => {
  const { fileName, base64, replaceExisting, sheetName, headerRow } = req.body || {};
  if (!base64) {
    return res.status(400).json({ error: 'Archivo inválido: falta contenido base64' });
  }

  let workbook;
  try {
    const buffer = Buffer.from(String(base64), 'base64');
    workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  } catch (err) {
    return res.status(400).json({ error: 'No se pudo leer el archivo Excel (.xlsx)' });
  }

  const defaultSheetName = workbook.SheetNames.includes('Bonanza')
    ? 'Bonanza'
    : (workbook.SheetNames && workbook.SheetNames[0]);
  const targetSheetName = String(sheetName || defaultSheetName || '').trim();
  if (!targetSheetName) {
    return res.status(400).json({ error: 'El Excel no contiene hojas' });
  }
  const sheet = workbook.Sheets[targetSheetName];
  if (!sheet) {
    return res.status(400).json({ error: `No existe la hoja "${targetSheetName}" en el archivo` });
  }
  const effectiveHeaderRow = Number(headerRow) > 0
    ? Number(headerRow)
    : (targetSheetName === 'Bonanza' ? 6 : 1);
  const rows = rowsFromSheetWithHeaderRow(sheet, effectiveHeaderRow);
  const parsedUnits = normalizeExcelRows(rows);

  const current = readOwnerServicesData();
  const baseUnits = Boolean(replaceExisting) ? [] : [...current.units];
  let created = 0;
  let updated = 0;

  parsedUnits.forEach((incoming) => {
    const idx = baseUnits.findIndex((u) => String(u.unitId || '').trim() === incoming.unitId);
    if (idx >= 0) {
      baseUnits[idx] = { ...baseUnits[idx], ...incoming };
      updated += 1;
    } else {
      baseUnits.push(incoming);
      created += 1;
    }
  });

  const next = { ...current, units: baseUnits };
  saveOwnerServicesData(next);

  return res.json({
    ok: true,
    fileName: String(fileName || ''),
    sheetName: targetSheetName,
    headerRow: effectiveHeaderRow,
    totalRows: rows.length,
    importedRows: parsedUnits.length,
    created,
    updated,
    skipped: Math.max(0, rows.length - parsedUnits.length)
  });
});

app.delete('/api/owner-services/units', (req, res) => {
  const current = readOwnerServicesData();
  const removed = Array.isArray(current.units) ? current.units.length : 0;
  const next = { ...current, units: [] };
  saveOwnerServicesData(next);
  return res.json({ ok: true, removed });
});

app.get('/format/:id', (req, res) => {
  const id = req.params.id;
  const format = formats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const data = fs.existsSync(DATA_PATH) ? JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) : {};
  const html = renderTemplate(format.file, data, {
    baseUrl,
    showControls: !req.query.print,
    downloadUrl: `/format/${id}/pdf`,
    bodyClass: 'interactive'
  });
  res.send(html);
});

app.get('/format/:id/pdf', async (req, res) => {
  const id = req.params.id;
  const format = formats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const puppeteer = require('puppeteer');
  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const data = fs.existsSync(DATA_PATH) ? JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) : {};
  const html = renderTemplate(format.file, data, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format/${id}/pdf`,
    bodyClass: 'print'
  });

  const browser = await puppeteer.launch({ headless: 'new' });
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const pdf = await page.pdf({
      format: 'Letter',
      printBackground: true,
      margin: { top: '0in', right: '0in', bottom: '0in', left: '0in' }
    });

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename=${format.file.replace('.html', '')}.pdf`
    });
    res.send(pdf);
  } finally {
    await browser.close();
  }
});

app.post('/format/:id/pdf', async (req, res) => {
  const id = req.params.id;
  const format = formats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const data = { ...(req.body || {}) };
  if (id === '10' || id === '19' || id === '21' || id === '35') {
    // Replace empty text fields with N/A
    Object.keys(data).forEach((k) => {
      const v = data[k];
      if (typeof v === 'string' && v.trim() === '') data[k] = 'N/A';
    });
  }

  const saved = persistSubmission(id, format.name, data);
  log(`Submission guardado: ${saved.id} formato=${id}`);

  const puppeteer = require('puppeteer');
  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const html = renderTemplate(format.file, data, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format/${id}/pdf`,
    bodyClass: 'print'
  });

  const browser = await puppeteer.launch({ headless: 'new' });
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const pdfOpts = {
      format: 'letter',
      printBackground: true,
      margin: { top: '0in', right: '0in', bottom: '0.6in', left: '0in' }
    };
    // FR-VEN-19 footer is rendered in HTML to match FR-VEN-10 layout
    const pdf = await page.pdf(pdfOpts);

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename=${format.file.replace('.html', '')}.pdf`
    });
    res.send(pdf);
  } finally {
    await browser.close();
  }
});

const server = app.listen(PORT, HOST, () => {
  log(`Servidor listo en http://${HOST}:${PORT}`);
});
server.on('error', (err) => {
  log(`Error al iniciar servidor: ${err && err.stack ? err.stack : err}`);
});
