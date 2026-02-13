const express = require('express');
const fs = require('fs');
const path = require('path');
const os = require('os');
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

const TEMPLATE_DIR = path.join(__dirname, 'templates');
const PUBLIC_DIR = path.join(__dirname, 'public');
const DATA_PATH = path.join(__dirname, 'data', 'sample.json');
const DATA_DIR = path.join(__dirname, 'data');
const SUBMISSIONS_PATH = path.join(DATA_DIR, 'submissions.json');
const OWNER_SERVICES_PATH = path.join(DATA_DIR, 'owner-services.json');
const FLOOR_JSON_DIR = path.join(DATA_DIR, 'plano-ventas-floors');

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
    .tag{display:inline-block; font-size:12px; font-weight:700; background:var(--accent); padding:4px 8px; border-radius:999px; margin-bottom:10px;}
    .name{font-size:20px; margin:0 0 8px;}
    .desc{margin:0; color:var(--muted);}
    .top{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;}
    .user{font-size:13px;color:#5f5f5f;}
    .logout{display:inline-block;padding:8px 10px;border:1px solid #bdb8a9;border-radius:10px;background:#fff;color:#111;text-decoration:none;font-size:13px;font-weight:600;}
  </style></head><body>
    <div class="wrap">
      <div class="top">
        <div>
          <h1>Backend SIMCA</h1>
          <p class="sub">Panel principal de módulos.</p>
        </div>
        <div style="text-align:right;">
          <div class="user">${String(req.user && req.user.email || '')}</div>
          <a class="logout" href="/logout">Cerrar sesión</a>
        </div>
      </div>
      <div class="grid">
        <a class="card" href="/generador-faes">
          <span class="tag">Módulo</span>
          <h2 class="name">Generador FAES</h2>
          <p class="desc">Gestión y generación de documentos FAES.</p>
        </a>
        <a class="card" href="/plds">
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
        <button type="button" class="back" onclick="history.back()">Regresar</button>
      </div>
      <div class="grid">
        <a class="card" href="/plano-interactivo">
          <span class="tag">A</span>
          <h2 class="name">Plano Interactivo</h2>
          <p class="desc">Marcado de unidades por piso y guardado base.</p>
        </a>
        <a class="card" href="/plano-ventas">
          <span class="tag">B</span>
          <h2 class="name">Plano Ventas (Editor)</h2>
          <p class="desc">Editor visual completo de PDF, tabla y estado de unidades.</p>
        </a>
        <a class="card" href="/plano-descargar">
          <span class="tag">C</span>
          <h2 class="name">Plano Descargar</h2>
          <p class="desc">Descarga rápida del PDF con carga automática de datos.</p>
        </a>
      </div>
    </div>
  </body></html>`);
});

app.get('/plano-descargar', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'plano-descargar.html'));
});

app.get('/api/plano-ventas/default-excel', (req, res) => {
  const filePath = path.join(os.homedir(), 'Downloads', 'INVENTARIOMAESTROWIX.xls');
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({
      error: 'Archivo no encontrado',
      expectedPath: filePath
    });
  }
  res.sendFile(filePath);
});

app.get('/api/plano-ventas/default-json', (req, res) => {
  const downloadsDir = path.join(os.homedir(), 'Downloads');
  let filePath = "";

  const requestName = typeof req.query.name === 'string' ? req.query.name.trim() : '';
  const safeName = requestName && !requestName.includes('/') && !requestName.includes('\\') ? requestName : '';
  if (safeName) {
    filePath = path.join(FLOOR_JSON_DIR, safeName);
    if (!fs.existsSync(filePath)) filePath = path.join(downloadsDir, safeName);
  } else {
    filePath = path.join(FLOOR_JSON_DIR, 'unidades-marcadas.JSON');
    if (!fs.existsSync(filePath)) filePath = path.join(FLOOR_JSON_DIR, 'unidades-marcadas.json');
  }

  if (!fs.existsSync(filePath)) {
    try {
      const files = fs.readdirSync(FLOOR_JSON_DIR);
      const preferred = files.find((name) => name.toLowerCase() === String(path.basename(filePath)).toLowerCase());
      if (preferred) filePath = path.join(FLOOR_JSON_DIR, preferred);
    } catch {}
  }

  if (!fs.existsSync(filePath) && !safeName) {
    try {
      const files = fs.readdirSync(FLOOR_JSON_DIR)
        .filter((name) => /^unidades-marcadas(?:\s*\(\d+\))?\.json$/i.test(name))
        .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
      if (files[0]) filePath = path.join(FLOOR_JSON_DIR, files[0]);
    } catch {}
  }

  if (!fs.existsSync(filePath)) {
    try {
      const files = fs.readdirSync(downloadsDir);
      const preferred = files.find((name) => name.toLowerCase() === String(path.basename(filePath)).toLowerCase());
      if (preferred) filePath = path.join(downloadsDir, preferred);
    } catch {}
  }

  if (!fs.existsSync(filePath) && !safeName) {
    try {
      const files = fs.readdirSync(downloadsDir)
        .filter((name) => /^unidades-marcadas(?:\s*\(\d+\))?\.json$/i.test(name))
        .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
      if (files[0]) filePath = path.join(downloadsDir, files[0]);
    } catch {}
  }

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({
      error: 'Archivo no encontrado',
      expectedPath: safeName ? path.join(FLOOR_JSON_DIR, safeName) : path.join(FLOOR_JSON_DIR, 'unidades-marcadas.json')
    });
  }

  res.sendFile(filePath);
});

app.get('/api/plano-ventas/default-json-files', (req, res) => {
  const downloadsDir = path.join(os.homedir(), 'Downloads');
  try {
    const primaryFiles = fs.readdirSync(FLOOR_JSON_DIR)
      .filter((name) => /^unidades-marcadas(?:\s*\(\d+\))?\.json$/i.test(name))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

    if (primaryFiles.length) {
      return res.json({ files: primaryFiles, sourceDir: FLOOR_JSON_DIR });
    }

    const files = fs.readdirSync(downloadsDir)
      .filter((name) => /^unidades-marcadas(?:\s*\(\d+\))?\.json$/i.test(name))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
    res.json({ files, sourceDir: downloadsDir });
  } catch (err) {
    res.status(500).json({ error: 'No se pudo leer carpeta de JSONs', details: err.message });
  }
});

app.get('/api/plano-ventas/default-json-merged', (req, res) => {
  const downloadsDir = path.join(os.homedir(), 'Downloads');
  try {
    let sourceDir = FLOOR_JSON_DIR;
    let files = fs.readdirSync(FLOOR_JSON_DIR)
      .filter((name) => /^unidades-marcadas(?:\s*\(\d+\))?\.json$/i.test(name))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));

    if (!files.length) {
      sourceDir = downloadsDir;
      files = fs.readdirSync(downloadsDir)
        .filter((name) => /^unidades-marcadas(?:\s*\(\d+\))?\.json$/i.test(name))
        .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }));
    }

    const floors = [];
    const loadedFiles = [];
    files.forEach((name) => {
      const filePath = path.join(sourceDir, name);
      try {
        const raw = fs.readFileSync(filePath, 'utf-8');
        const parsed = JSON.parse(raw);
        const payloadFloors = Array.isArray(parsed)
          ? parsed
          : (parsed && Array.isArray(parsed.floors) ? parsed.floors : (parsed && parsed.imageDataUrl ? [parsed] : []));
        if (payloadFloors.length) {
          floors.push(...payloadFloors);
          loadedFiles.push(name);
        }
      } catch {}
    });

    res.json({ floors, loadedFiles, sourceDir });
  } catch (err) {
    res.status(500).json({ error: 'No se pudo leer JSONs de carpeta control', details: err.message });
  }
});

app.post('/api/plano-ventas/render-pdf', async (req, res) => {
  const html = typeof req.body?.html === 'string' ? req.body.html : '';
  if (!html || html.length < 20) {
    return res.status(400).json({ error: 'HTML inválido para generar PDF.' });
  }

  let browser = null;
  try {
    browser = await puppeteer.launch({
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    await page.emulateMediaType('print');

    const pdfBuffer = await page.pdf({
      printBackground: true,
      preferCSSPageSize: true
    });

    const utfName = 'CEIBA INVENTARIO ／ INVENTORY.pdf';
    const fallbackName = 'CEIBA INVENTARIO - INVENTORY.pdf';
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${fallbackName}"; filename*=UTF-8''${encodeURIComponent(utfName)}`
    );
    return res.send(pdfBuffer);
  } catch (err) {
    log(`Error en /api/plano-ventas/render-pdf: ${err && err.stack ? err.stack : err}`);
    return res.status(500).json({ error: 'No se pudo generar el PDF en el servidor.' });
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
