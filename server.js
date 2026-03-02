const express = require('express');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execSync } = require('child_process');
process.env.PUPPETEER_CACHE_DIR = path.join(__dirname, '.cache', 'puppeteer');
const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const QRCode = require('qrcode');
const session = require('express-session');
const passport = require('passport');
const GoogleStrategy = require('passport-google-oauth20').Strategy;
const MicrosoftStrategy = require('passport-microsoft').Strategy;
const { Pool } = require('pg');
const app = express();
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '0.0.0.0';
const APP_BASE_URL = process.env.APP_BASE_URL || `http://localhost:${PORT}`;
const APP_BASE_URL_NORMALIZED = String(APP_BASE_URL).replace(/\/+$/, '');
const APP_BUILD_ID = String(
  process.env.RENDER_GIT_COMMIT
  || process.env.RAILWAY_GIT_COMMIT_SHA
  || process.env.VERCEL_GIT_COMMIT_SHA
  || `${Date.now()}-${process.pid}`
);
const SESSION_SECRET = process.env.SESSION_SECRET || 'dev-session-secret-change-me';
const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '';
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || '';
const GOOGLE_AUTH_READY = Boolean(GOOGLE_CLIENT_ID && GOOGLE_CLIENT_SECRET);
const MICROSOFT_CLIENT_ID = process.env.MICROSOFT_CLIENT_ID || '';
const MICROSOFT_CLIENT_SECRET = process.env.MICROSOFT_CLIENT_SECRET || '';
const MICROSOFT_TENANT_ID = process.env.MICROSOFT_TENANT_ID || 'common';
const MICROSOFT_AUTH_READY = Boolean(MICROSOFT_CLIENT_ID && MICROSOFT_CLIENT_SECRET);
const LEAD_TOKEN_SECRET = process.env.LEAD_TOKEN_SECRET || SESSION_SECRET;
const VICEROY_PRESENT_TOKEN = String(process.env.VICEROY_PRESENT_TOKEN || '').trim();
const LOCAL_NO_AUTH = String(process.env.LOCAL_NO_AUTH || '') === '1';
const ALLOWED_DOMAIN = String(process.env.ALLOWED_DOMAIN || 'simca.mx').toLowerCase();
const GERENTE_EMAIL = String(process.env.GERENTE_EMAIL || 'martin@simca.mx').toLowerCase();
const EXTRA_ALLOWED_EMAILS = new Set(
  String(process.env.ALLOWED_EMAILS || '')
    .split(',')
    .map((item) => String(item || '').trim().toLowerCase())
    .filter(Boolean)
);
const DATABASE_URL = String(process.env.DATABASE_URL || '').trim();
const USE_WHISPERLIST_DB = Boolean(DATABASE_URL);
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
const REPO_DATA_DIR = path.join(__dirname, 'data');
const DATA_DIR = path.resolve(String(process.env.APP_DATA_DIR || REPO_DATA_DIR));
const DATA_PATH = path.join(DATA_DIR, 'sample.json');
const ROI_MASTER_CSV_PATH = path.join(DATA_DIR, 'roi-master.csv');
const SUBMISSIONS_PATH = path.join(DATA_DIR, 'submissions.json');
const OWNER_SERVICES_PATH = path.join(DATA_DIR, 'owner-services.json');
const WHISPERLIST_JSON_PATH = path.join(DATA_DIR, 'viceroy-whisperlist.json');
const WHISPERLIST_EXCEL_PATH = path.join(DATA_DIR, 'VICEROY WHISPERLIST.xlsx');
const VICEROY_PILOTO_CONFIG_NAME = 'viceroy-tipologias.json';
const FLOOR_JSON_DIR = path.join(DATA_DIR, 'plano-ventas-floors');
const DEVELOPMENTS_DIR = path.join(DATA_DIR, 'developments');
const SEED_DEVELOPMENTS_DIR = path.join(__dirname, 'seed-data', 'developments');
const DEFAULT_DEVELOPMENT_SLUG = 'ceiba';
const FLOOR_JSON_FILE_RE = /^unidades-marcadas(?:\s*\(\d+\))?\.json$/i;
const FLOOR_MAPPED_JSON_FILE_RE = /^imagen-mapeada-.*\.json$/i;
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
  { slug: 'solar-mt', name: 'SOLAR MT' },
  { slug: 'viceroy-piloto', name: 'VICEROY PILOTO' }
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
const nationalFormats = {
  '35': { name: 'FR-VEN-35 Aviso de Privacidad (Nacional PF)', file: 'format-35-nacional.html' },
  '21': { name: 'FR-VEN-21 Consulta de Listas de Personas Bloqueadas', file: 'format-21.html' },
  '26': { name: 'FR-VEN-26 Origen de los Recursos (Nacional PF)', file: 'format-26-nacional.html' },
  '19': { name: 'FR-VEN-19 Beneficiario Controlador PF (Nacional)', file: 'format-19-nacional.html' },
  '10': { name: 'FR-VEN-10 Identificación del Cliente (Nacional PF)', file: 'format-10-nacional.html' }
};
const nationalMoralFormats = {
  '35': { name: 'FR-VEN-35 Aviso de Privacidad (Nacional PM)', file: 'format-35-nacional.html' },
  '21': { name: 'FR-VEN-21 Consulta de Listas de Personas Bloqueadas', file: 'format-21.html' },
  '26': { name: 'FR-VEN-26 Origen de los Recursos (Nacional PM)', file: 'format-26-nacional.html' },
  '19': { name: 'FR-VEN-19 Beneficiario Controlador PM (Nacional)', file: 'format-19-nacional-moral.html' },
  '10': { name: 'FR-VEN-10 Identificación del Cliente (Nacional PM)', file: 'format-10-nacional-moral.html' }
};
const WHISPERLIST_CANALES = ['SIMCA', 'RELATED'];
const WHISPERLIST_TIPOS_VENTA = ['EXTERNO', 'INTERNO', 'MERITO PROPIO'];
const WHISPERLIST_RECAMARAS = ['1', '2', '3', '2PH', '3PH'];

const whisperlistPool = USE_WHISPERLIST_DB
  ? new Pool({
    connectionString: DATABASE_URL,
    ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false
  })
  : null;
let whisperlistStorageReady = false;

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
if (LOCAL_NO_AUTH) {
  app.use((req, _res, next) => {
    req.user = {
      id: 'local-dev-user',
      email: GERENTE_EMAIL,
      name: 'Local Dev'
    };
    req.isAuthenticated = () => true;
    next();
  });
}
app.use('/assets', express.static(path.join(PUBLIC_DIR, 'assets')));
app.use('/plds-static', express.static(path.join(PUBLIC_DIR, 'plds')));

passport.serializeUser((user, done) => done(null, user));
passport.deserializeUser((user, done) => done(null, user));

if (GOOGLE_AUTH_READY) {
  passport.use(new GoogleStrategy(
    {
      clientID: GOOGLE_CLIENT_ID,
      clientSecret: GOOGLE_CLIENT_SECRET,
      callbackURL: `${APP_BASE_URL_NORMALIZED}/auth/google/callback`
    },
    (accessToken, refreshToken, profile, done) => {
      (async () => {
        const rawEmail = profile && profile.emails && profile.emails[0] ? profile.emails[0].value : '';
        const email = String(rawEmail || '').trim().toLowerCase();
        if (!(await isAllowedLoginEmail(email))) {
          return done(null, false, { message: 'Correo no autorizado' });
        }
        return done(null, {
          id: profile.id,
          email,
          name: profile.displayName || email
        });
      })().catch((err) => done(err));
    }
  ));
}

if (MICROSOFT_AUTH_READY) {
  passport.use(new MicrosoftStrategy(
    {
      clientID: MICROSOFT_CLIENT_ID,
      clientSecret: MICROSOFT_CLIENT_SECRET,
      callbackURL: `${APP_BASE_URL_NORMALIZED}/auth/microsoft/callback`,
      scope: ['user.read'],
      tenant: MICROSOFT_TENANT_ID,
      addUPNAsEmail: true
    },
    (accessToken, refreshToken, profile, done) => {
      (async () => {
        const profileEmail = profile && Array.isArray(profile.emails) && profile.emails[0]
          ? profile.emails[0].value
          : '';
        const upnEmail = profile && profile._json ? (profile._json.userPrincipalName || profile._json.mail || '') : '';
        const email = String(profileEmail || upnEmail || '').trim().toLowerCase();
        if (!(await isAllowedLoginEmail(email))) {
          return done(null, false, { message: 'Correo no autorizado' });
        }
        return done(null, {
          id: profile.id,
          email,
          name: profile.displayName || email
        });
      })().catch((err) => done(err));
    }
  ));
}

function requireAuth(req, res, next) {
  if (LOCAL_NO_AUTH) return next();
  if (req.isAuthenticated && req.isAuthenticated()) return next();
  if (String(req.path || '').startsWith('/api/')) {
    return res.status(401).json({ error: 'No autenticado' });
  }
  return res.redirect('/login');
}

function isInternalUserEmail(email) {
  const normalized = String(email || '').trim().toLowerCase();
  if (!normalized) return false;
  if (normalized === GERENTE_EMAIL) return true;
  return normalized.endsWith(`@${ALLOWED_DOMAIN}`);
}

function requireInternalUser(req, res, next) {
  if (LOCAL_NO_AUTH) return next();
  if (!(req.isAuthenticated && req.isAuthenticated())) {
    if (String(req.path || '').startsWith('/api/')) {
      return res.status(401).json({ error: 'No autenticado' });
    }
    return res.redirect('/login');
  }
  const email = String(req.user && req.user.email || '').toLowerCase();
  if (isInternalUserEmail(email)) return next();
  if (String(req.path || '').startsWith('/api/')) {
    return res.status(403).json({ error: 'Acceso restringido a usuarios internos' });
  }
  return res.redirect('/whisperlist');
}

function requireGerente(req, res, next) {
  if (LOCAL_NO_AUTH) return next();
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

function requireViceroyPresentAccess(req, res, next) {
  if (!VICEROY_PRESENT_TOKEN) return next();
  const token = String(req.query && req.query.token || '').trim();
  if (token && token === VICEROY_PRESENT_TOKEN) return next();
  return res.status(403).send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Acceso restringido</title>
  <style>
    body{font-family:Arial,sans-serif;margin:0;background:#f4f1e8;color:#1a1a1a;display:grid;place-items:center;min-height:100vh;padding:20px;}
    .card{width:min(620px,100%);background:#fff;border:1px solid #d8d1c1;border-radius:14px;padding:22px;}
    h1{margin:0 0 10px;font-size:28px;}
    p{margin:0 0 12px;color:#5f5f5f;line-height:1.4;}
  </style></head>
  <body>
    <section class="card">
      <h1>Acceso protegido</h1>
      <p>Este enlace de presentación requiere token.</p>
      <p>Abre la URL con <code>?token=TU_TOKEN</code>.</p>
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
  const copyIfMissing = (src, dst) => {
    try {
      if (!fs.existsSync(src) || fs.existsSync(dst)) return;
      const parent = path.dirname(dst);
      if (!fs.existsSync(parent)) fs.mkdirSync(parent, { recursive: true });
      fs.copyFileSync(src, dst);
    } catch {}
  };

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

  // Bootstraps Viceroy files in mutable storage (Render disk, etc.) once.
  const viceroyDir = path.join(DEVELOPMENTS_DIR, 'viceroy-piloto');
  const viceroyFloorDir = path.join(viceroyDir, 'plano-ventas-floors');
  if (!fs.existsSync(viceroyFloorDir)) fs.mkdirSync(viceroyFloorDir, { recursive: true });

  copyIfMissing(
    path.join(SEED_DEVELOPMENTS_DIR, 'viceroy-piloto', VICEROY_PILOTO_CONFIG_NAME),
    path.join(viceroyDir, VICEROY_PILOTO_CONFIG_NAME)
  );
  copyIfMissing(
    path.join(SEED_DEVELOPMENTS_DIR, 'viceroy-piloto', 'plano-ventas-floors', 'imagen-mapeada-viceroy-piloto-2026-02-23T00-58-00-513Z.json'),
    path.join(viceroyFloorDir, 'imagen-mapeada-viceroy-piloto-2026-02-23T00-58-00-513Z.json')
  );
  copyIfMissing(
    path.join(REPO_DATA_DIR, 'developments', 'viceroy-piloto', 'INVENTARIOMAESTROWIX.xlsx'),
    path.join(viceroyDir, 'INVENTARIOMAESTROWIX.xlsx')
  );
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
  if (devSlug === 'viceroy-piloto') {
    const seedFloorDir = path.join(SEED_DEVELOPMENTS_DIR, devSlug, 'plano-ventas-floors');
    return [primary, FLOOR_JSON_DIR, seedFloorDir];
  }
  return [primary];
}

function getViceroyPilotoConfigPath() {
  return path.join(DEVELOPMENTS_DIR, 'viceroy-piloto', VICEROY_PILOTO_CONFIG_NAME);
}

function getViceroyPilotoSeedConfigPath() {
  return path.join(SEED_DEVELOPMENTS_DIR, 'viceroy-piloto', VICEROY_PILOTO_CONFIG_NAME);
}

function defaultViceroyPilotoConfig() {
  return {
    tipologias: [],
    unitTipologiaMap: {},
    unitRecamarasMap: {},
    mapFloorOrder: [],
    presentationLayout: {
      showBrand: false,
      showLevel: false,
      overlayX: 18,
      overlayY: 18,
      overlayWidth: 360,
      overlayBg: 'rgba(255,255,255,0.88)',
      overlayBorder: '#cdd2c6',
      tableHeaderBg: '#ecece4',
      tableHeaderText: '#666a75',
      tableText: '#31353e',
      tableFontFamily: '"Akkurat", "Segoe UI", Tahoma, sans-serif',
      tableHeaderFontFamily: '"Akkurat", "Segoe UI", Tahoma, sans-serif',
      tableFontSize: 12,
      tableHeaderFontSize: 12,
      tableTextAlign: 'left',
      tableHeaderAlign: 'left',
      overlayRefWidth: 1500,
      overlayRefHeight: 900,
      colWUnidad: 24,
      colWRec: 12,
      colWVista: 20,
      colWM2: 12,
      colWPrecio: 32
    },
    pages: [],
    updatedAt: new Date().toISOString()
  };
}

function normalizeViceroyPresentationLayout(raw) {
  const source = raw && typeof raw === 'object' ? raw : {};
  const toNum = (value, fallback, min, max) => {
    const n = Number(value);
    if (!Number.isFinite(n)) return fallback;
    return Math.max(min, Math.min(max, n));
  };
  const toColor = (value, fallback) => {
    const txt = String(value || '').trim();
    if (!txt) return fallback;
    return txt;
  };
  const toText = (value, fallback, maxLen = 160) => {
    const txt = String(value || '').trim();
    if (!txt) return fallback;
    return txt.slice(0, maxLen);
  };
  const toAlign = (value, fallback = 'left') => {
    const txt = String(value || '').trim().toLowerCase();
    if (txt === 'left' || txt === 'center' || txt === 'right') return txt;
    return fallback;
  };
  return {
    showBrand: Boolean(source.showBrand),
    showLevel: Boolean(source.showLevel),
    overlayX: toNum(source.overlayX, 18, 0, 5000),
    overlayY: toNum(source.overlayY, 18, 0, 5000),
    overlayWidth: toNum(source.overlayWidth, 360, 220, 1800),
    overlayBg: toColor(source.overlayBg, 'rgba(255,255,255,0.88)'),
    overlayBorder: toColor(source.overlayBorder, '#cdd2c6'),
    tableHeaderBg: toColor(source.tableHeaderBg, '#ecece4'),
    tableHeaderText: toColor(source.tableHeaderText, '#666a75'),
    tableText: toColor(source.tableText, '#31353e'),
    tableFontFamily: toText(source.tableFontFamily, '"Akkurat", "Segoe UI", Tahoma, sans-serif'),
    tableHeaderFontFamily: toText(source.tableHeaderFontFamily, '"Akkurat", "Segoe UI", Tahoma, sans-serif'),
    tableFontSize: toNum(source.tableFontSize, 12, 9, 30),
    tableHeaderFontSize: toNum(source.tableHeaderFontSize, 12, 9, 30),
    tableTextAlign: toAlign(source.tableTextAlign, 'left'),
    tableHeaderAlign: toAlign(source.tableHeaderAlign, 'left'),
    overlayRefWidth: toNum(source.overlayRefWidth, 1500, 200, 8000),
    overlayRefHeight: toNum(source.overlayRefHeight, 900, 200, 8000),
    colWUnidad: toNum(source.colWUnidad, 24, 5, 80),
    colWRec: toNum(source.colWRec, 12, 5, 80),
    colWVista: toNum(source.colWVista, 20, 5, 80),
    colWM2: toNum(source.colWM2, 12, 5, 80),
    colWPrecio: toNum(source.colWPrecio, 32, 5, 80)
  };
}

function normalizeViceroyPilotoConfig(raw) {
  const source = raw && typeof raw === 'object' ? raw : {};
  const tipologias = Array.isArray(source.tipologias)
    ? source.tipologias.map((item, index) => {
      const id = String(item && item.id || '').trim();
      const name = String(item && item.name || '').trim();
      const recamaras = String(item && item.recamaras || '').trim().toUpperCase().replace(/\s+/g, '');
      const imageUrl = String(item && item.imageUrl || '').trim();
      if (!id && !name) return null;
      return {
        id: id || `TIPO-${index + 1}`,
        name: name || id || `TIPO-${index + 1}`,
        recamaras,
        imageUrl
      };
    }).filter(Boolean)
    : [];
  const unitTipologiaMap = {};
  if (source.unitTipologiaMap && typeof source.unitTipologiaMap === 'object') {
    Object.entries(source.unitTipologiaMap).forEach(([unit, tipologia]) => {
      const unitKey = String(unit || '').trim().toUpperCase();
      const tipologiaValue = String(tipologia || '').trim();
      if (!unitKey || !tipologiaValue) return;
      unitTipologiaMap[unitKey] = tipologiaValue;
    });
  }
  const unitRecamarasMap = {};
  if (source.unitRecamarasMap && typeof source.unitRecamarasMap === 'object') {
    Object.entries(source.unitRecamarasMap).forEach(([unit, rec]) => {
      const unitKey = String(unit || '').trim().toUpperCase();
      const recValue = String(rec || '').trim().toUpperCase().replace(/\s+/g, '');
      if (!unitKey || !recValue) return;
      unitRecamarasMap[unitKey] = recValue;
    });
  }
  const pages = Array.isArray(source.pages)
    ? source.pages.map((item, index) => {
      const id = String(item && item.id || '').trim();
      const title = String(item && item.title || '').trim();
      const imageUrl = String(item && item.imageUrl || '').trim();
      const targetType = String(item && item.targetType || '').trim().toLowerCase();
      const targetValue = String(item && item.targetValue || '').trim();
      if (!id && !title && !imageUrl) return null;
      return {
        id: id || `PAGE-${index + 1}`,
        title: title || id || `PAGE ${index + 1}`,
        imageUrl,
        targetType: ['tipologia', 'recamaras', 'all'].includes(targetType) ? targetType : '',
        targetValue
      };
    }).filter(Boolean)
    : [];
  const mapFloorOrder = Array.isArray(source.mapFloorOrder)
    ? source.mapFloorOrder
      .map((item) => String(item || '').trim())
      .filter(Boolean)
    : [];
  const presentationLayout = normalizeViceroyPresentationLayout(source.presentationLayout);
  return {
    tipologias,
    unitTipologiaMap,
    unitRecamarasMap,
    mapFloorOrder,
    presentationLayout,
    pages,
    layoutSource: String(source.layoutSource || '').trim().toLowerCase() === 'custom' ? 'custom' : '',
    updatedAt: new Date().toISOString()
  };
}

function getViceroyInventoryCachePath() {
  return path.join(DEVELOPMENTS_DIR, 'viceroy-piloto', 'inventory-cache.json');
}

function readViceroyInventoryCache() {
  const cachePath = getViceroyInventoryCachePath();
  if (!fs.existsSync(cachePath)) return { rows: [], fileName: '' };
  try {
    const raw = JSON.parse(fs.readFileSync(cachePath, 'utf-8'));
    const rows = Array.isArray(raw.rows) ? raw.rows : [];
    const fileName = String(raw.fileName || '').trim();
    return { rows, fileName };
  } catch {
    return { rows: [], fileName: '' };
  }
}

function writeViceroyInventoryCache(rows, fileName) {
  const cachePath = getViceroyInventoryCachePath();
  const dir = path.dirname(cachePath);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  const payload = {
    fileName: String(fileName || '').trim(),
    rows: Array.isArray(rows) ? rows : [],
    updatedAt: new Date().toISOString()
  };
  fs.writeFileSync(cachePath, JSON.stringify(payload, null, 2), 'utf-8');
}

function readViceroyPilotoConfig() {
  const filePath = getViceroyPilotoConfigPath();
  let seedConfig = null;
  const seedPath = getViceroyPilotoSeedConfigPath();
  if (fs.existsSync(seedPath)) {
    try {
      const rawSeed = JSON.parse(fs.readFileSync(seedPath, 'utf-8'));
      seedConfig = normalizeViceroyPilotoConfig(rawSeed);
    } catch {}
  }
  if (!fs.existsSync(filePath)) return seedConfig || defaultViceroyPilotoConfig();
  try {
    const raw = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    const normalized = normalizeViceroyPilotoConfig(raw);
    const defaultLayout = defaultViceroyPilotoConfig().presentationLayout;
    const normalizedLayout = normalized && normalized.presentationLayout && typeof normalized.presentationLayout === 'object'
      ? normalized.presentationLayout
      : defaultLayout;
    const isDefaultLikeLayout = JSON.stringify(normalizedLayout) === JSON.stringify(defaultLayout);

    const runtimeLayoutSource = String(normalized.layoutSource || '').trim().toLowerCase();
    // Seed layout rules:
    // 1) If runtime is default-like, prefer local seed.
    // 2) If runtime has no explicit custom marker, seed still wins (one-time migration from local).
    if (seedConfig && (isDefaultLikeLayout || runtimeLayoutSource !== 'custom')) {
      const merged = {
        ...normalized,
        presentationLayout: {
          ...(normalized.presentationLayout || {}),
          ...(seedConfig.presentationLayout || {})
        }
      };
      return normalizeViceroyPilotoConfig(merged);
    }
    // Self-heal legacy/incomplete files so presentationLayout is always persisted.
    const rawLayout = raw && raw.presentationLayout && typeof raw.presentationLayout === 'object' ? raw.presentationLayout : null;
    const hasLayoutKeys = rawLayout && (
      Object.prototype.hasOwnProperty.call(rawLayout, 'overlayX') ||
      Object.prototype.hasOwnProperty.call(rawLayout, 'overlayY') ||
      Object.prototype.hasOwnProperty.call(rawLayout, 'overlayWidth')
    );
    if (!hasLayoutKeys) {
      try {
        fs.writeFileSync(filePath, JSON.stringify(normalized, null, 2), 'utf-8');
      } catch {}
    }
    return normalized;
  } catch {
    return defaultViceroyPilotoConfig();
  }
}

function writeViceroyPilotoConfig(payload) {
  const filePath = getViceroyPilotoConfigPath();
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  const normalized = normalizeViceroyPilotoConfig(payload);
  fs.writeFileSync(filePath, JSON.stringify(normalized, null, 2), 'utf-8');
  return normalized;
}

function updateViceroyPilotoPresentationLayout(layoutPayload) {
  const current = readViceroyPilotoConfig();
  const next = {
    ...current,
    layoutSource: 'custom',
    presentationLayout: normalizeViceroyPresentationLayout(layoutPayload)
  };
  return writeViceroyPilotoConfig(next);
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
      '--font-render-hinting=medium'
    ]
  };
  return puppeteer.launch(launchOptions);
}

let sharedPdfBrowser = null;
async function getSharedPdfBrowser() {
  if (sharedPdfBrowser) {
    try {
      if (sharedPdfBrowser.isConnected()) return sharedPdfBrowser;
    } catch {}
  }
  try {
    sharedPdfBrowser = await launchPdfBrowser();
  } catch (launchErr) {
    const installed = installChromeIfMissing(launchErr);
    if (!installed) throw launchErr;
    sharedPdfBrowser = await launchPdfBrowser();
  }
  sharedPdfBrowser.on('disconnected', () => {
    sharedPdfBrowser = null;
  });
  return sharedPdfBrowser;
}

async function buildPdfBufferWithBrowser(browser, html, pdfOptions = {}) {
  const page = await browser.newPage();
  try {
    page.setDefaultTimeout(120000);
    page.setDefaultNavigationTimeout(120000);
    await page.setViewport({ width: 1600, height: 2200, deviceScaleFactor: 2 });
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 120000 });
    await page.emulateMediaType('print');
    try {
      const hasReadyMarker = await page.evaluate(() => Object.prototype.hasOwnProperty.call(window, '__pdfReady'));
      if (hasReadyMarker) {
        await page.waitForFunction(() => window.__pdfReady === true, { timeout: 8000 });
      }
    } catch {
      // If marker check fails, continue and generate PDF.
    }
    return await page.pdf({
      printBackground: true,
      preferCSSPageSize: true,
      ...pdfOptions
    });
  } finally {
    try { await page.close(); } catch {}
  }
}

function isRetryablePdfError(err) {
  const msg = String((err && err.message) || '');
  return /Target closed|Session closed|browser has disconnected|Protocol error/i.test(msg);
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

function normalizeWhisperlistKey(raw) {
  return String(raw || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function normalizeWhisperlistCanal(raw) {
  const value = String(raw || '').trim().toUpperCase();
  return WHISPERLIST_CANALES.includes(value) ? value : '';
}

function normalizeWhisperlistTipoVenta(raw) {
  const value = String(raw || '').trim().toUpperCase().replace(/\s+/g, ' ');
  if (value === 'MERITO PROPIO' || value === 'MÉRITO PROPIO') return 'MERITO PROPIO';
  if (WHISPERLIST_TIPOS_VENTA.includes(value)) return value;
  return '';
}

function normalizeWhisperlistRecamaras(raw) {
  const value = String(raw || '').trim().toUpperCase().replace(/\s+/g, '');
  return WHISPERLIST_RECAMARAS.includes(value) ? value : '';
}

function normalizeWhisperlistPersonText(raw) {
  return String(raw || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}

function normalizeWhisperlistAsesor(raw) {
  const normalized = normalizeWhisperlistPersonText(raw);
  const key = normalized
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  if (key === 'ABIGAIL PEREZ') return 'ABIGAIL ZARATE';
  return normalized;
}

function normalizeClientEmail(raw) {
  return String(raw || '').trim().toLowerCase();
}

function normalizeClientPhone(raw) {
  return String(raw || '').trim().replace(/\s+/g, ' ');
}

function encodeBase64Url(raw) {
  return Buffer.from(String(raw || ''), 'utf-8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function decodeBase64Url(value) {
  const normalized = String(value || '').replace(/-/g, '+').replace(/_/g, '/');
  const padded = normalized + '==='.slice((normalized.length + 3) % 4);
  return Buffer.from(padded, 'base64').toString('utf-8');
}

function signLeadPayload(payloadText) {
  return crypto
    .createHmac('sha256', LEAD_TOKEN_SECRET)
    .update(String(payloadText || ''))
    .digest('hex');
}

function createLeadToken(payload) {
  const payloadText = JSON.stringify(payload || {});
  const encoded = encodeBase64Url(payloadText);
  const signature = signLeadPayload(encoded);
  return `${encoded}.${signature}`;
}

function verifyLeadToken(token) {
  const rawToken = String(token || '').trim();
  const parts = rawToken.split('.');
  if (parts.length !== 2) return null;
  const [encoded, signature] = parts;
  const expected = signLeadPayload(encoded);
  if (signature !== expected) return null;
  try {
    const payload = JSON.parse(decodeBase64Url(encoded));
    const asesor = normalizeWhisperlistAsesor(payload.asesor);
    const correo = String(payload.correo || '').trim().toLowerCase();
    if (!asesor || !correo) return null;
    return { asesor, correo };
  } catch {
    return null;
  }
}

function normalizeWhisperlistRow(rawRow, fallbackId) {
  const normalized = {};
  Object.entries(rawRow || {}).forEach(([key, value]) => {
    normalized[normalizeWhisperlistKey(key)] = value;
  });

  const asesor = normalizeWhisperlistAsesor(normalized.asesor);
  const correo = String(normalized.correo || '').trim().toLowerCase();
  if (!asesor || !correo) return null;

  const generatedId = `${Date.now()}-${Math.random().toString(36).slice(2, 10)}-${Math.random().toString(36).slice(2, 6)}`;
  const providedId = String(normalized.id || '').trim();
  return {
    id: providedId || String(fallbackId || generatedId),
    asesor,
    correo,
    canal: normalizeWhisperlistCanal(normalized.canal),
    tipoVenta: normalizeWhisperlistTipoVenta(normalized.tipo_de_venta || normalized.tipodeventa),
    nombreCliente: normalizeWhisperlistPersonText(normalized.nombre_cliente || normalized.nombrecliente),
    recamaras: normalizeWhisperlistRecamaras(normalized.recamaras || normalized.recamara),
    clientEmail: normalizeClientEmail(normalized.correo_cliente || normalized.email_cliente || normalized.client_email),
    clientPhone: normalizeClientPhone(normalized.telefono_cliente || normalized.telefono || normalized.client_phone),
    updatedAt: new Date().toISOString()
  };
}

function parseWhisperlistWorkbook(workbook, requestedSheetName) {
  const fallbackSheet = workbook.SheetNames.includes('Hoja1')
    ? 'Hoja1'
    : (workbook.SheetNames[0] || '');
  const sheetName = String(requestedSheetName || fallbackSheet || '').trim();
  if (!sheetName) {
    return { sheetName: '', rows: [] };
  }
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    return { sheetName: '', rows: [] };
  }
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  const rows = rawRows
    .map((item) => normalizeWhisperlistRow(item))
    .filter(Boolean);

  return { sheetName, rows };
}

async function ensureWhisperlistDbSchema() {
  if (!whisperlistPool) return;
  await whisperlistPool.query(`
    CREATE TABLE IF NOT EXISTS whisperlist_rows (
      id TEXT PRIMARY KEY,
      asesor TEXT NOT NULL,
      correo TEXT NOT NULL,
      canal TEXT NOT NULL DEFAULT '',
      tipo_venta TEXT NOT NULL DEFAULT '',
      nombre_cliente TEXT NOT NULL DEFAULT '',
      recamaras TEXT NOT NULL DEFAULT '',
      client_email TEXT NOT NULL DEFAULT '',
      client_phone TEXT NOT NULL DEFAULT '',
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )
  `);
  await whisperlistPool.query(`ALTER TABLE whisperlist_rows ADD COLUMN IF NOT EXISTS recamaras TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE whisperlist_rows ADD COLUMN IF NOT EXISTS client_email TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE whisperlist_rows ADD COLUMN IF NOT EXISTS client_phone TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`
    CREATE TABLE IF NOT EXISTS whisperlist_meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    )
  `);
}

async function readWhisperlistData() {
  if (!whisperlistStorageReady) await ensureWhisperlistStorageReady();
  if (!whisperlistPool) {
    const raw = readJson(WHISPERLIST_JSON_PATH, { rows: [], updatedAt: null, sourceFile: '' });
    const rows = sortWhisperlistRows(Array.isArray(raw.rows) ? raw.rows : []);
    return {
      rows,
      updatedAt: raw.updatedAt || null,
      sourceFile: String(raw.sourceFile || '')
    };
  }

  const [rowsRes, metaRes] = await Promise.all([
    whisperlistPool.query(`
      SELECT id, asesor, correo, canal, tipo_venta, nombre_cliente, recamaras, client_email, client_phone, updated_at
      FROM whisperlist_rows
      ORDER BY updated_at DESC, id ASC
    `),
    whisperlistPool.query(`
      SELECT key, value
      FROM whisperlist_meta
      WHERE key IN ('sourceFile', 'updatedAt')
    `)
  ]);

  const meta = {};
  metaRes.rows.forEach((item) => {
    meta[String(item.key || '')] = String(item.value || '');
  });

  return {
    rows: sortWhisperlistRows(rowsRes.rows.map((row) => ({
      id: String(row.id || ''),
      asesor: String(row.asesor || ''),
      correo: String(row.correo || '').toLowerCase(),
      canal: String(row.canal || ''),
      tipoVenta: String(row.tipo_venta || ''),
      nombreCliente: String(row.nombre_cliente || ''),
      recamaras: String(row.recamaras || ''),
      clientEmail: String(row.client_email || ''),
      clientPhone: String(row.client_phone || ''),
      updatedAt: row.updated_at ? new Date(row.updated_at).toISOString() : new Date().toISOString()
    }))),
    updatedAt: meta.updatedAt || null,
    sourceFile: meta.sourceFile || ''
  };
}

async function saveWhisperlistRows(rows, sourceFile) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const normalizedRows = safeRows.map((row) => ({
    ...row,
    asesor: normalizeWhisperlistAsesor(row.asesor),
    nombreCliente: normalizeWhisperlistPersonText(row.nombreCliente),
    correo: String(row.correo || '').trim().toLowerCase(),
    clientEmail: normalizeClientEmail(row.clientEmail),
    clientPhone: normalizeClientPhone(row.clientPhone),
    recamaras: normalizeWhisperlistRecamaras(row.recamaras),
    canal: normalizeWhisperlistCanal(row.canal),
    tipoVenta: normalizeWhisperlistTipoVenta(row.tipoVenta)
  }));
  const updatedAt = new Date().toISOString();
  if (!whisperlistPool) {
    writeJson(WHISPERLIST_JSON_PATH, {
      rows: normalizedRows,
      updatedAt,
      sourceFile: String(sourceFile || '')
    });
    return;
  }

  const client = await whisperlistPool.connect();
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM whisperlist_rows');
    for (const row of normalizedRows) {
      await client.query(
        `INSERT INTO whisperlist_rows (id, asesor, correo, canal, tipo_venta, nombre_cliente, recamaras, client_email, client_phone, updated_at)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)`,
        [
          String(row.id || `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`),
          normalizeWhisperlistAsesor(row.asesor),
          String(row.correo || '').trim().toLowerCase(),
          String(row.canal || '').trim(),
          String(row.tipoVenta || '').trim(),
          normalizeWhisperlistPersonText(row.nombreCliente),
          normalizeWhisperlistRecamaras(row.recamaras),
          normalizeClientEmail(row.clientEmail),
          normalizeClientPhone(row.clientPhone),
          String(row.updatedAt || updatedAt)
        ]
      );
    }
    await client.query(
      `INSERT INTO whisperlist_meta (key, value)
       VALUES ('sourceFile', $1)
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
      [String(sourceFile || '')]
    );
    await client.query(
      `INSERT INTO whisperlist_meta (key, value)
       VALUES ('updatedAt', $1)
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
      [updatedAt]
    );
    await client.query('COMMIT');
  } catch (err) {
    await client.query('ROLLBACK');
    throw err;
  } finally {
    client.release();
  }
}

async function whisperlistAllowedEmails() {
  const data = await readWhisperlistData();
  const emails = new Set();
  data.rows.forEach((row) => {
    const email = String(row && row.correo || '').trim().toLowerCase();
    if (email) emails.add(email);
  });
  return emails;
}

async function isAllowedLoginEmail(email) {
  const normalized = String(email || '').trim().toLowerCase();
  if (!normalized) return false;
  if (normalized === GERENTE_EMAIL) return true;
  if (normalized.endsWith(`@${ALLOWED_DOMAIN}`)) return true;
  if (EXTRA_ALLOWED_EMAILS.has(normalized)) return true;
  if ((await whisperlistAllowedEmails()).has(normalized)) return true;
  return false;
}

async function generatePdfWithSharedBrowser(html, pdfOptions = {}) {
  try {
    const browser = await getSharedPdfBrowser();
    return await buildPdfBufferWithBrowser(browser, html, pdfOptions);
  } catch (firstErr) {
    if (!isRetryablePdfError(firstErr)) throw firstErr;
    try {
      if (sharedPdfBrowser) await sharedPdfBrowser.close();
    } catch {}
    sharedPdfBrowser = null;
    const retryBrowser = await getSharedPdfBrowser();
    return await buildPdfBufferWithBrowser(retryBrowser, html, pdfOptions);
  }
}

function getWhisperlistSellers(rows) {
  const items = Array.isArray(rows) ? rows : [];
  const byEmail = new Map();
  items.forEach((row) => {
    const correo = String(row && row.correo || '').trim().toLowerCase();
    const asesor = normalizeWhisperlistAsesor(row && row.asesor || '');
    if (!correo || !asesor) return;
    if (!byEmail.has(correo)) {
      byEmail.set(correo, { asesor, correo });
    }
  });
  return Array.from(byEmail.values()).sort((a, b) => a.asesor.localeCompare(b.asesor));
}

function sortWhisperlistRows(rows) {
  const list = Array.isArray(rows) ? [...rows] : [];
  return list.sort((a, b) => {
    const asesorA = normalizeWhisperlistAsesor(a && a.asesor || '');
    const asesorB = normalizeWhisperlistAsesor(b && b.asesor || '');
    const byAsesor = asesorA.localeCompare(asesorB, 'es');
    if (byAsesor !== 0) return byAsesor;
    const clienteA = normalizeWhisperlistPersonText(a && a.nombreCliente || '');
    const clienteB = normalizeWhisperlistPersonText(b && b.nombreCliente || '');
    return clienteA.localeCompare(clienteB, 'es');
  });
}

function whisperlistExportRows(rows) {
  const items = Array.isArray(rows) ? rows : [];
  return items.map((row) => ({
    ASESOR: normalizeWhisperlistAsesor(row.asesor),
    CORREO_ASESOR: String(row.correo || '').trim().toLowerCase(),
    CANAL: normalizeWhisperlistCanal(row.canal),
    TIPO_DE_VENTA: normalizeWhisperlistTipoVenta(row.tipoVenta),
    NOMBRE_CLIENTE: normalizeWhisperlistPersonText(row.nombreCliente),
    RECAMARAS: normalizeWhisperlistRecamaras(row.recamaras),
    CORREO_CLIENTE: normalizeClientEmail(row.clientEmail),
    TELEFONO_CLIENTE: normalizeClientPhone(row.clientPhone),
    UPDATED_AT: String(row.updatedAt || '')
  }));
}

async function seedWhisperlistFromExcelIfNeeded() {
  if (!fs.existsSync(WHISPERLIST_EXCEL_PATH)) return;
  if (!whisperlistPool && fs.existsSync(WHISPERLIST_JSON_PATH)) return;
  try {
    if (whisperlistPool) {
      const countRes = await whisperlistPool.query('SELECT COUNT(*)::int AS total FROM whisperlist_rows');
      const count = Number(countRes.rows[0] && countRes.rows[0].total || 0);
      if (count > 0) return;
    }
    const workbook = XLSX.readFile(WHISPERLIST_EXCEL_PATH, { cellDates: true });
    const parsed = parseWhisperlistWorkbook(workbook);
    await saveWhisperlistRows(parsed.rows, path.basename(WHISPERLIST_EXCEL_PATH));
    log(`Whisperlist inicializada con ${parsed.rows.length} filas`);
  } catch (err) {
    log(`No se pudo inicializar whisperlist: ${err && err.message ? err.message : err}`);
  }
}

async function ensureWhisperlistStorageReady() {
  if (whisperlistStorageReady) return;
  log(`Whisperlist storage: ${whisperlistPool ? 'postgres' : 'json-file'}`);
  if (whisperlistPool) await ensureWhisperlistDbSchema();
  await seedWhisperlistFromExcelIfNeeded();
  whisperlistStorageReady = true;
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

const VICEROY_PILOTO_COLUMN_ALIASES = {
  development: ['development', 'desarrollo', 'proyecto'],
  unit: ['unidad', 'departamento', 'depto', 'unit', 'unit_id', 'unitid', 'no'],
  recamaras: ['recamaras', 'recamaras_', 'rec', 'bedrooms', 'beds', 'habitaciones'],
  tipologia: ['tipologia', 'tipologia_', 'typology', 'tipo', 'tipo_unidad', 'type'],
  view: ['vista', 'view', 'vistas'],
  m2: ['m2', 'metros2', 'metros_cuadrados', 'metros cuadrados', 'm²', 'area_m2', 'area'],
  sqft: ['sqft', 'ft2', 'pies2', 'pies_cuadrados', 'square_feet', 'area_sqft'],
  price: ['precio_final', 'precio final', 'precio', 'precio_venta', 'venta', 'sale_price', 'price'],
  status: ['estatus', 'estado', 'status', 'disponibilidad', 'availability', 'inventario']
};

function viceroyPilotoTargetDevKeys() {
  return new Set(['viceroy', 'viceroy_piloto', 'viceroy_residences', 'viceroy_playa_del_carmen']);
}

function pickValueByAliases(record, aliases) {
  for (const alias of aliases) {
    if (Object.prototype.hasOwnProperty.call(record, alias)) {
      const value = record[alias];
      if (value !== undefined && value !== null && String(value).trim() !== '') return value;
    }
  }
  return '';
}

function normalizeViceroyRowStatus(rawValue) {
  const value = normalizeHeaderKey(rawValue);
  if (!value) return 'disponible';
  if (value.includes('vendid') || value.includes('sold') || value.includes('agotad') || value.includes('no_disponible')) {
    return 'vendida';
  }
  if (value.includes('apartad') || value.includes('reservad')) {
    return 'vendida';
  }
  return 'disponible';
}

function normalizeViceroyInventoryRow(rawRow) {
  const row = {};
  Object.entries(rawRow || {}).forEach(([key, value]) => {
    row[normalizeHeaderKey(key)] = value;
  });
  const unit = String(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.unit) || '').trim();
  if (!unit) return null;
  const development = String(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.development) || '').trim();
  const rawType = String(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.tipologia) || '').trim();
  let recamaras = String(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.recamaras) || '').trim().toUpperCase().replace(/\s+/g, '');
  if (!recamaras && rawType) {
    const typeKey = normalizeHeaderKey(rawType);
    if ((typeKey.includes('ph') || typeKey.includes('pent')) && typeKey.includes('3')) recamaras = '3PH';
    else if ((typeKey.includes('ph') || typeKey.includes('pent')) && typeKey.includes('2')) recamaras = '2PH';
    else if (typeKey.includes('3')) recamaras = '3B';
    else if (typeKey.includes('2')) recamaras = '2B';
    else if (typeKey.includes('1')) recamaras = '1B';
  }
  const tipologia = rawType;
  const view = String(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.view) || '').trim();
  const m2Raw = pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.m2);
  const sqftRaw = pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.sqft);
  const priceRaw = pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.price);
  const status = normalizeViceroyRowStatus(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.status));
  return {
    development,
    unidad: unit,
    recamaras,
    tipologia,
    vista: view,
    m2: m2Raw === '' ? '' : m2Raw,
    sqft: sqftRaw === '' ? '' : sqftRaw,
    price: priceRaw === '' ? '' : priceRaw,
    status
  };
}

function normalizeViceroyRawBedroom(value) {
  const key = normalizeHeaderKey(value);
  if (!key) return '';
  if ((key.includes('ph') || key.includes('pent')) && (key.includes('3') || key.includes('three') || key.includes('tres'))) return '3PH';
  if ((key.includes('ph') || key.includes('pent')) && (key.includes('2') || key.includes('two') || key.includes('dos'))) return '2PH';
  if (key.includes('3') || key.includes('three') || key.includes('tres')) return '3B';
  if (key.includes('2') || key.includes('two') || key.includes('dos')) return '2B';
  if (key.includes('1') || key.includes('one') || key.includes('uno') || key.includes('una') || key === 'un') return '1B';
  return '';
}

function rowDataByIndex(row, index) {
  const value = Array.isArray(row) ? row[index] : '';
  return value === undefined || value === null ? '' : value;
}

function parseViceroyRowsByFixedColumns(sheet) {
  const out = [];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
  if (!Array.isArray(matrix) || !matrix.length) return out;

  matrix.forEach((row) => {
    const unidad = String(rowDataByIndex(row, 0) || '').trim(); // A
    if (!unidad) return;
    const unitKey = normalizeHeaderKey(unidad);
    if (!unitKey) return;
    if (['unidad', 'unit', 'departamento', 'depto', 'no'].includes(unitKey)) return;

    const tipologia = String(rowDataByIndex(row, 3) || '').trim(); // D
    const vista = String(rowDataByIndex(row, 4) || '').trim(); // E
    const m2 = rowDataByIndex(row, 14); // O
    const recRaw = rowDataByIndex(row, 15); // P
    const price = rowDataByIndex(row, 19); // T

    out.push({
      development: '',
      unidad,
      recamaras: normalizeViceroyRawBedroom(recRaw),
      tipologia,
      vista,
      m2: m2 === '' ? '' : m2,
      sqft: '',
      price: price === '' ? '' : price,
      status: 'disponible'
    });
  });

  return out;
}

function pickRicherViceroyRow(current, next) {
  if (!current) return next;
  if (!next) return current;
  const currentScore = ['recamaras', 'tipologia', 'vista', 'm2', 'sqft', 'price'].reduce((acc, key) => {
    return acc + (String(current[key] || '').trim() ? 1 : 0);
  }, 0);
  const nextScore = ['recamaras', 'tipologia', 'vista', 'm2', 'sqft', 'price'].reduce((acc, key) => {
    return acc + (String(next[key] || '').trim() ? 1 : 0);
  }, 0);
  return nextScore >= currentScore ? next : current;
}

function dedupeViceroyRows(rows) {
  const byUnit = new Map();
  rows.forEach((row) => {
    const key = String(row && row.unidad || '').trim().toUpperCase();
    if (!key) return;
    byUnit.set(key, pickRicherViceroyRow(byUnit.get(key), row));
  });
  return Array.from(byUnit.values());
}

function parseViceroyInventoryFile(filePath) {
  const ext = path.extname(String(filePath || '')).toLowerCase();
  if (ext === '.csv') {
    const rawText = fs.readFileSync(filePath, 'utf-8');
    const workbook = XLSX.read(rawText, { type: 'string' });
    const firstSheetName = workbook.SheetNames[0];
    const firstSheet = firstSheetName ? workbook.Sheets[firstSheetName] : null;
    if (!firstSheet) return [];
    const headerRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' })
      .map((row) => normalizeViceroyInventoryRow(row))
      .filter(Boolean);
    const indexRows = parseViceroyRowsByFixedColumns(firstSheet);
    return dedupeViceroyRows([...headerRows, ...indexRows]);
  }

  const workbook = XLSX.readFile(filePath);
  const out = [];
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    rows.forEach((row) => {
      const normalized = normalizeViceroyInventoryRow(row);
      if (normalized) out.push(normalized);
    });
    out.push(...parseViceroyRowsByFixedColumns(sheet));
  });
  return dedupeViceroyRows(out);
}

function filterViceroyInventoryRows(rows) {
  const targetDevKeys = viceroyPilotoTargetDevKeys();
  let pendingDev = '';
  return rows.filter((row) => {
    const rowDevKey = normalizeHeaderKey(row.development);
    if (rowDevKey) pendingDev = rowDevKey;
    const effectiveDevKey = rowDevKey || pendingDev;
    if (!effectiveDevKey) return true;
    return targetDevKeys.has(effectiveDevKey);
  });
}

function resolveInventoryCandidatesByDevSlug(devSlug) {
  const devDirs = [];
  const primary = path.join(DEVELOPMENTS_DIR, devSlug);
  devDirs.push(primary);
  const repoDir = path.join(REPO_DATA_DIR, 'developments', devSlug);
  if (repoDir !== primary) devDirs.push(repoDir);
  const files = [];
  devDirs.forEach((devDir) => {
    try {
      fs.readdirSync(devDir).forEach((name) => {
        if (!/\.(xls|xlsx|csv)$/i.test(name)) return;
        const fullPath = path.join(devDir, name);
        let mtimeMs = 0;
        try { mtimeMs = fs.statSync(fullPath).mtimeMs || 0; } catch {}
        files.push({ name, fullPath, mtimeMs });
      });
    } catch {}
  });
  const canonicalNames = new Set([
    'inventariomaestrowix.xls',
    'inventariomaestrowix.xlsx',
    'inventariomaestrowix.csv'
  ]);
  const canonical = files
    .filter((f) => canonicalNames.has(String(f.name || '').toLowerCase()))
    .sort((a, b) => b.mtimeMs - a.mtimeMs);
  const latest = files
    .filter((f) => !canonicalNames.has(String(f.name || '').toLowerCase()))
    .sort((a, b) => b.mtimeMs - a.mtimeMs);
  return [...canonical, ...latest];
}

function readMergedFloorsByDevelopment(devSlug) {
  const floorDirs = getDevelopmentFloorSearchDirs(devSlug);
  const entries = [];
  floorDirs.forEach((dir) => {
    try {
      fs.readdirSync(dir)
        .filter((name) => {
          const lower = String(name || '').toLowerCase();
          if (!lower.endsWith('.json')) return false;
          if (/^version-.*\.json$/i.test(name)) return true;
          if (FLOOR_MAPPED_JSON_FILE_RE.test(name)) return true;
          return FLOOR_JSON_FILE_RE.test(name);
        })
        .forEach((name) => {
          const fullPath = path.join(dir, name);
          let mtimeMs = 0;
          try { mtimeMs = fs.statSync(fullPath).mtimeMs || 0; } catch {}
          entries.push({ name, fullPath, mtimeMs });
        });
    } catch {}
  });

  const unique = [];
  const seen = new Set();
  entries
    .sort((a, b) => b.mtimeMs - a.mtimeMs || a.name.localeCompare(b.name))
    .forEach((entry) => {
      const key = String(entry.name || '').toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      unique.push(entry);
    });

  const floors = [];
  const loadedFiles = [];
  let filesToLoad = [];
  if (devSlug === 'viceroy-piloto') {
    const mapped = unique.filter((entry) => FLOOR_MAPPED_JSON_FILE_RE.test(entry.name));
    const canonical = unique.filter((entry) => FLOOR_JSON_FILE_RE.test(entry.name));
    const rest = unique.filter((entry) => !FLOOR_MAPPED_JSON_FILE_RE.test(entry.name) && !FLOOR_JSON_FILE_RE.test(entry.name));
    filesToLoad = [...mapped, ...canonical, ...rest];
  } else {
    filesToLoad = [...unique];
  }

  for (const entry of filesToLoad) {
    try {
      const raw = JSON.parse(fs.readFileSync(entry.fullPath, 'utf-8'));
      const payloadFloors = Array.isArray(raw)
        ? raw
        : (raw && Array.isArray(raw.floors) ? raw.floors : (raw && raw.imageDataUrl ? [raw] : []));
      if (!payloadFloors.length) continue;
      floors.push(...payloadFloors);
      loadedFiles.push(path.basename(entry.fullPath));
      if (devSlug === 'viceroy-piloto' && floors.length) break;
    } catch {}
  }
  return { floors, loadedFiles };
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

function withPldsFooterCode(data, scope, id) {
  const merged = { ...(data || {}) };
  if (id === '26' && !merged.orFooterCode) {
    merged.orFooterCode = scope === 'NM' ? 'FR-VEN-26 OR PM N' : 'FR-VEN-26 OR PF N';
  }
  return merged;
}

app.get('/login', (req, res) => {
  if (req.isAuthenticated && req.isAuthenticated()) return res.redirect('/');
  const googleReady = GOOGLE_AUTH_READY;
  const microsoftReady = MICROSOFT_AUTH_READY;
  const anyReady = googleReady || microsoftReady;
  const error = req.query && req.query.error === 'domain'
    ? 'Correo no autorizado. Usa un correo permitido.'
    : '';
  const providersHtml = [
    googleReady ? '<a class="btn" href="/auth/google">Entrar con Google</a>' : '',
    microsoftReady ? '<a class="btn" href="/auth/microsoft">Entrar con Microsoft</a>' : ''
  ].filter(Boolean).join('<span style="display:inline-block;width:8px"></span>');
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
      <p>Inicia sesión con Google o Microsoft. Acceso para cuentas <strong>@${ALLOWED_DOMAIN}</strong> o correos invitados autorizados.</p>
      ${anyReady ? providersHtml : '<p class="warn">Falta configurar OAuth. Variables requeridas: GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET o MICROSOFT_CLIENT_ID/MICROSOFT_CLIENT_SECRET.</p>'}
      ${error ? `<p class="warn">${error}</p>` : ''}
      <p class="muted">Gerente ventas permitido: ${GERENTE_EMAIL}</p>
    </div>
  </body></html>`);
});

if (GOOGLE_AUTH_READY) {
  app.get('/auth/google', passport.authenticate('google', {
    scope: ['profile', 'email']
  }));

  app.get('/auth/google/callback',
    passport.authenticate('google', { failureRedirect: '/login?error=domain' }),
    (req, res) => res.redirect('/')
  );
} else {
  app.get('/auth/google', (req, res) => res.redirect('/login'));
  app.get('/auth/google/callback', (req, res) => res.redirect('/login'));
}

if (MICROSOFT_AUTH_READY) {
  app.get('/auth/microsoft', passport.authenticate('microsoft', {
    prompt: 'select_account'
  }));

  app.get('/auth/microsoft/callback',
    passport.authenticate('microsoft', { failureRedirect: '/login?error=domain' }),
    (req, res) => res.redirect('/')
  );
} else {
  app.get('/auth/microsoft', (req, res) => res.redirect('/login'));
  app.get('/auth/microsoft/callback', (req, res) => res.redirect('/login'));
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
  const isInternalUser = isInternalUserEmail(currentEmail);
  if (!isInternalUser) {
    return res.redirect('/whisperlist');
  }
  const isGerente = currentEmail === GERENTE_EMAIL;
  const showViceroyPilotoCard = currentEmail === 'martin@simca.mx';
  const ownerServicesCard = isGerente ? `
        <a class="card" href="/owner-services">
          <span class="tag">Módulo</span>
          <h2 class="name">Owner Services</h2>
          <p class="desc">Prioriza entregas y coordina obra, jurídico y finanzas.</p>
        </a>` : '';
  const viceroyPilotoCard = showViceroyPilotoCard ? `
        <a class="card" href="/viceroy-piloto">
          <span class="tag">Módulo</span>
          <h2 class="name">VICEROY PILOTO</h2>
          <p class="desc">Flujo visual de recámaras, tipologías y plano por piso.</p>
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
        ${viceroyPilotoCard}
        <a class="card" href="/whisperlist">
          <span class="tag">Módulo</span>
          <h2 class="name">Viceroy Whisperlist</h2>
          <p class="desc">Todos ven la tabla, cada asesor edita solo sus filas.</p>
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

app.use('/legacy', requireInternalUser);
app.use('/generador-faes', requireInternalUser);
app.use('/plds', requireInternalUser);
app.use('/generador-roi', requireInternalUser);
app.use('/form', requireInternalUser);
app.use('/form-nacional', requireInternalUser);
app.use('/form-nacional-moral', requireInternalUser);
app.use('/format', requireInternalUser);
app.use('/format-nacional', requireInternalUser);
app.use('/format-nacional-moral', requireInternalUser);
app.use('/submissions', requireInternalUser);
app.use('/api/plds', requireInternalUser);
app.use('/api/roi', requireInternalUser);
app.use('/viceroy-piloto', requireInternalUser);
app.use('/api/viceroy-piloto', (req, res, next) => {
  if (req.path === '/public-data') return next();
  return requireInternalUser(req, res, next);
});
app.use('/whisperlist/qr', requireGerente);
app.use('/whisperlist', requireAuth);
app.use('/api/whisperlist', requireAuth);
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
  res.redirect('/form-nacional');
});

app.get('/plds/cliente-nacional-persona-moral', (req, res) => {
  res.redirect('/form-nacional-moral');
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

app.get('/whisperlist', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'whisperlist.html'));
});

app.get('/api/app-version', (req, res) => {
  return res.json({
    ok: true,
    buildId: APP_BUILD_ID
  });
});

app.get('/whisperlist/qr', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'whisperlist-qr.html'));
});

app.get('/api/whisperlist', async (req, res) => {
  try {
    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const isGerente = currentEmail === GERENTE_EMAIL;
    const data = await readWhisperlistData();
    const rows = data.rows.map((row) => ({
      ...row,
      asesor: normalizeWhisperlistAsesor(row.asesor),
      nombreCliente: normalizeWhisperlistPersonText(row.nombreCliente),
      recamaras: normalizeWhisperlistRecamaras(row.recamaras),
      clientEmail: (isGerente || String(row.correo || '').toLowerCase() === currentEmail)
        ? normalizeClientEmail(row.clientEmail)
        : '',
      clientPhone: (isGerente || String(row.correo || '').toLowerCase() === currentEmail)
        ? normalizeClientPhone(row.clientPhone)
        : '',
      canEdit: isGerente || String(row.correo || '').toLowerCase() === currentEmail,
      canViewClientData: isGerente || String(row.correo || '').toLowerCase() === currentEmail
    }));
    res.json({
      ok: true,
      currentEmail,
      isGerente,
      updatedAt: data.updatedAt,
      sourceFile: data.sourceFile,
      rows
    });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo leer whisperlist' });
  }
});

app.get('/api/whisperlist/qr-codes', requireGerente, async (req, res) => {
  try {
    const data = await readWhisperlistData();
    const sellers = getWhisperlistSellers(data.rows);
    const items = await Promise.all(sellers.map(async (seller) => {
      const token = createLeadToken({ asesor: seller.asesor, correo: seller.correo });
      const leadUrl = `${APP_BASE_URL_NORMALIZED}/lead?t=${encodeURIComponent(token)}`;
      const qrDataUrl = await QRCode.toDataURL(leadUrl, { margin: 1, width: 280 });
      return {
        ...seller,
        leadUrl,
        qrDataUrl
      };
    }));
    return res.json({ ok: true, items });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudieron generar QR' });
  }
});

app.get('/api/whisperlist/export.xlsx', requireGerente, async (req, res) => {
  try {
    const data = await readWhisperlistData();
    const rows = whisperlistExportRows(data.rows);
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Whisperlist');
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    const dateTag = new Date().toISOString().slice(0, 10);
    const fileName = `whisperlist-backup-${dateTag}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    return res.send(buffer);
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo exportar whisperlist' });
  }
});

app.patch('/api/whisperlist/rows/:id', async (req, res) => {
  try {
    const rowId = String(req.params.id || '').trim();
    if (!rowId) return res.status(400).json({ error: 'id de fila inválido' });

    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const isGerente = currentEmail === GERENTE_EMAIL;
    const data = await readWhisperlistData();
    const index = data.rows.findIndex((row) => String(row.id || '') === rowId);
    if (index < 0) return res.status(404).json({ error: 'Fila no encontrada' });

    const target = data.rows[index];
    const ownerEmail = String(target.correo || '').trim().toLowerCase();
    if (!isGerente && ownerEmail !== currentEmail) {
      return res.status(403).json({ error: 'Solo puedes editar filas asignadas a tu correo' });
    }

    const body = req.body || {};
    const nextRow = {
      ...target,
      canal: normalizeWhisperlistCanal(body.canal !== undefined ? body.canal : target.canal),
      tipoVenta: normalizeWhisperlistTipoVenta(body.tipoVenta !== undefined ? body.tipoVenta : target.tipoVenta),
      nombreCliente: normalizeWhisperlistPersonText(body.nombreCliente !== undefined ? body.nombreCliente : target.nombreCliente),
      recamaras: normalizeWhisperlistRecamaras(body.recamaras !== undefined ? body.recamaras : target.recamaras),
      clientEmail: normalizeClientEmail(body.clientEmail !== undefined ? body.clientEmail : target.clientEmail),
      clientPhone: normalizeClientPhone(body.clientPhone !== undefined ? body.clientPhone : target.clientPhone),
      updatedAt: new Date().toISOString()
    };
    data.rows[index] = nextRow;
    await saveWhisperlistRows(data.rows, data.sourceFile || path.basename(WHISPERLIST_EXCEL_PATH));
    return res.json({ ok: true, row: nextRow });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo guardar fila' });
  }
});

app.post('/api/whisperlist/rows', async (req, res) => {
  try {
    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const fallbackName = String(req.user && req.user.name || '').trim();
    const data = await readWhisperlistData();
    const existing = data.rows.find((row) => String(row.correo || '').trim().toLowerCase() === currentEmail);
    const asesor = normalizeWhisperlistAsesor(existing && existing.asesor || fallbackName || currentEmail.split('@')[0] || '');

    const body = req.body || {};
    const nombreCliente = normalizeWhisperlistPersonText(body.nombreCliente);
    if (!nombreCliente) return res.status(400).json({ error: 'nombreCliente es obligatorio' });

    const newRow = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      asesor,
      correo: currentEmail,
      canal: normalizeWhisperlistCanal(body.canal),
      tipoVenta: normalizeWhisperlistTipoVenta(body.tipoVenta),
      nombreCliente,
      recamaras: normalizeWhisperlistRecamaras(body.recamaras),
      clientEmail: normalizeClientEmail(body.clientEmail),
      clientPhone: normalizeClientPhone(body.clientPhone),
      updatedAt: new Date().toISOString()
    };
    data.rows.push(newRow);
    await saveWhisperlistRows(data.rows, data.sourceFile || path.basename(WHISPERLIST_EXCEL_PATH));
    return res.status(201).json({ ok: true, row: newRow });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo agregar fila' });
  }
});

app.delete('/api/whisperlist/rows/:id', async (req, res) => {
  try {
    const rowId = String(req.params.id || '').trim();
    if (!rowId) return res.status(400).json({ error: 'id de fila inválido' });

    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const isGerente = currentEmail === GERENTE_EMAIL;
    const data = await readWhisperlistData();
    const index = data.rows.findIndex((row) => String(row.id || '') === rowId);
    if (index < 0) return res.status(404).json({ error: 'Fila no encontrada' });

    const target = data.rows[index];
    const ownerEmail = String(target.correo || '').trim().toLowerCase();
    if (!isGerente && ownerEmail !== currentEmail) {
      return res.status(403).json({ error: 'Solo puedes eliminar filas asignadas a tu correo' });
    }

    data.rows.splice(index, 1);
    await saveWhisperlistRows(data.rows, data.sourceFile || path.basename(WHISPERLIST_EXCEL_PATH));
    return res.json({ ok: true });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo eliminar fila' });
  }
});

app.get('/lead', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'lead-form.html'));
});

app.post('/api/lead/submit', async (req, res) => {
  try {
    const token = String(req.body && req.body.token || '').trim();
    const decoded = verifyLeadToken(token);
    if (!decoded) return res.status(400).json({ error: 'Token inválido' });

    const nombreCliente = normalizeWhisperlistPersonText(req.body && req.body.nombreCliente);
    if (!nombreCliente) return res.status(400).json({ error: 'nombreCliente es obligatorio' });
    const clientEmail = normalizeClientEmail(req.body && req.body.clientEmail);
    const clientPhone = normalizeClientPhone(req.body && req.body.clientPhone);

    const newRow = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      asesor: decoded.asesor,
      correo: decoded.correo,
      canal: 'RELATED',
      tipoVenta: '',
      nombreCliente,
      recamaras: normalizeWhisperlistRecamaras(req.body && req.body.recamaras),
      clientEmail,
      clientPhone,
      updatedAt: new Date().toISOString()
    };

    const data = await readWhisperlistData();
    data.rows.push(newRow);
    await saveWhisperlistRows(data.rows, data.sourceFile || path.basename(WHISPERLIST_EXCEL_PATH));
    return res.status(201).json({ ok: true });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo registrar el cliente' });
  }
});

app.post('/api/whisperlist/import-excel', requireGerente, async (req, res) => {
  const { fileName, base64, sheetName, replaceExisting } = req.body || {};
  if (!base64) {
    return res.status(400).json({ error: 'Archivo inválido: falta contenido base64' });
  }

  let workbook;
  try {
    const payload = String(base64).includes(',') ? String(base64).split(',').pop() : String(base64);
    const buffer = Buffer.from(payload || '', 'base64');
    workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  } catch (err) {
    return res.status(400).json({ error: 'No se pudo leer el archivo Excel (.xlsx)' });
  }

  const parsed = parseWhisperlistWorkbook(workbook, sheetName);
  if (!parsed.sheetName) {
    return res.status(400).json({ error: 'No se encontró una hoja válida en el Excel' });
  }

  const safeName = sanitizeExcelFileName(fileName) || 'whisperlist.xlsx';
  const current = await readWhisperlistData();
  const shouldReplace = Boolean(replaceExisting);
  let created = 0;
  let skipped = 0;
  const nextRows = shouldReplace ? [] : [...current.rows];
  const seen = new Set(
    nextRows.map((row) => [
      String(row.correo || '').trim().toLowerCase(),
      String(row.asesor || '').trim().toUpperCase(),
      String(row.nombreCliente || '').trim().toUpperCase(),
      String(row.canal || '').trim().toUpperCase(),
      String(row.tipoVenta || '').trim().toUpperCase()
    ].join('|'))
  );
  parsed.rows.forEach((row) => {
    const key = [
      String(row.correo || '').trim().toLowerCase(),
      String(row.asesor || '').trim().toUpperCase(),
      String(row.nombreCliente || '').trim().toUpperCase(),
      String(row.canal || '').trim().toUpperCase(),
      String(row.tipoVenta || '').trim().toUpperCase()
    ].join('|');
    if (seen.has(key)) {
      skipped += 1;
      return;
    }
    nextRows.push(row);
    seen.add(key);
    created += 1;
  });

  await saveWhisperlistRows(nextRows, safeName);
  return res.json({
    ok: true,
    fileName: safeName,
    sheetName: parsed.sheetName,
    importedRows: parsed.rows.length,
    mode: shouldReplace ? 'replace' : 'append',
    created,
    skipped,
    totalRows: nextRows.length
  });
});

app.get('/owner-services', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'owner-services.html'));
});

app.get('/viceroy-piloto-presentacion', requireViceroyPresentAccess, (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-piloto-presentacion.html'));
});

app.get('/api/viceroy-piloto/public-data', requireViceroyPresentAccess, (req, res) => {
  const devSlug = 'viceroy-piloto';
  const floorsData = readMergedFloorsByDevelopment(devSlug);
  const config = readViceroyPilotoConfig();
  const candidates = resolveInventoryCandidatesByDevSlug(devSlug);
  let inventoryRows = [];
  let inventoryFileName = '';
  if (candidates.length) {
    try {
      inventoryRows = filterViceroyInventoryRows(parseViceroyInventoryFile(candidates[0].fullPath));
      inventoryFileName = candidates[0].name;
    } catch {}
  }
  if (!inventoryRows.length) {
    const cached = readViceroyInventoryCache();
    inventoryRows = cached.rows;
    inventoryFileName = cached.fileName || inventoryFileName;
  }
  return res.json({
    ok: true,
    dev: devSlug,
    floors: floorsData.floors,
    loadedFiles: floorsData.loadedFiles,
    config,
    inventoryFileName,
    inventoryRows
  });
});

app.get('/viceroy-piloto', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-piloto.html'));
});

app.get('/api/viceroy-piloto/config', (req, res) => {
  const config = readViceroyPilotoConfig();
  return res.json({
    ok: true,
    config
  });
});

app.get('/api/viceroy-piloto/upload-access', (req, res) => {
  const email = String(req.user && req.user.email || '').trim().toLowerCase();
  const canUpload = LOCAL_NO_AUTH || email === GERENTE_EMAIL;
  return res.json({ ok: true, canUpload });
});

app.post('/api/viceroy-piloto/config', (req, res) => {
  const body = req.body || {};
  const payload = body.config && typeof body.config === 'object' ? body.config : body;
  const saved = writeViceroyPilotoConfig(payload);
  return res.json({
    ok: true,
    config: saved
  });
});

app.post('/api/viceroy-piloto/layout', (req, res) => {
  const body = req.body || {};
  const layoutPayload = body.presentationLayout && typeof body.presentationLayout === 'object'
    ? body.presentationLayout
    : body;
  const saved = updateViceroyPilotoPresentationLayout(layoutPayload);
  return res.json({
    ok: true,
    presentationLayout: saved.presentationLayout
  });
});

app.get('/api/viceroy-piloto/floors', (req, res) => {
  const devSlug = 'viceroy-piloto';
  const data = readMergedFloorsByDevelopment(devSlug);
  return res.json({
    ok: true,
    dev: devSlug,
    floors: data.floors,
    loadedFiles: data.loadedFiles
  });
});

app.get('/api/viceroy-piloto/inventory', (req, res) => {
  const devSlug = 'viceroy-piloto';
  const candidates = resolveInventoryCandidatesByDevSlug(devSlug);
  if (!candidates.length) {
    return res.status(404).json({
      error: 'No se encontró inventario para VICEROY PILOTO',
      expectedDir: path.join(DEVELOPMENTS_DIR, devSlug)
    });
  }
  const file = candidates[0];
  try {
    const rows = filterViceroyInventoryRows(parseViceroyInventoryFile(file.fullPath));
    if (rows.length) writeViceroyInventoryCache(rows, file.name);
    return res.json({
      ok: true,
      fileName: file.name,
      rows
    });
  } catch (err) {
    const cached = readViceroyInventoryCache();
    if (cached.rows.length) {
      return res.json({
        ok: true,
        fileName: cached.fileName || file.name,
        rows: cached.rows
      });
    }
    return res.status(500).json({
      error: 'No se pudo leer inventario',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.post('/api/viceroy-piloto/inventory', (req, res) => {
  const email = String(req.user && req.user.email || '').trim().toLowerCase();
  if (!LOCAL_NO_AUTH && email !== GERENTE_EMAIL) {
    return res.status(403).json({ error: 'Solo gerente puede subir inventario' });
  }
  const rawFileName = sanitizeExcelFileName(req.body && req.body.fileName);
  const base64Content = String(req.body && req.body.base64 || '');
  if (!rawFileName || !base64Content) {
    return res.status(400).json({ error: 'Archivo inventario inválido' });
  }
  const commaIndex = base64Content.indexOf(',');
  const payload = commaIndex >= 0 ? base64Content.slice(commaIndex + 1) : base64Content;
  const buffer = Buffer.from(payload, 'base64');
  if (!buffer.length) {
    return res.status(400).json({ error: 'Contenido vacío' });
  }
  const devDir = path.join(DEVELOPMENTS_DIR, 'viceroy-piloto');
  if (!fs.existsSync(devDir)) fs.mkdirSync(devDir, { recursive: true });
  const ext = (rawFileName.split('.').pop() || '').toLowerCase();
  const canonicalName = `INVENTARIOMAESTROWIX.${ext}`;
  const canonicalPath = path.join(devDir, canonicalName);
  fs.writeFileSync(canonicalPath, buffer);
  // Mirror write to repo data path when runtime storage points elsewhere.
  const repoDevDir = path.join(REPO_DATA_DIR, 'developments', 'viceroy-piloto');
  if (repoDevDir !== devDir) {
    try {
      if (!fs.existsSync(repoDevDir)) fs.mkdirSync(repoDevDir, { recursive: true });
      fs.writeFileSync(path.join(repoDevDir, canonicalName), buffer);
    } catch {}
  }
  // Verify file is readable right away so frontend does not show false-positive success.
  try {
    const persistedRows = filterViceroyInventoryRows(parseViceroyInventoryFile(canonicalPath));
    if (!persistedRows.length) {
      return res.status(500).json({ error: 'Se guardó archivo pero no se pudo leer inventario persistido' });
    }
    writeViceroyInventoryCache(persistedRows, canonicalName);
  } catch (err) {
    return res.status(500).json({
      error: 'Archivo guardado pero lectura de verificación falló',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
  return res.json({
    ok: true,
    fileName: canonicalName
  });
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
  const onlyInteractive = dev.slug === 'viceroy-piloto';
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
        ${onlyInteractive ? '' : `<a class="card" href="/gerente-ventas/${dev.slug}/plano-ventas?dev=${dev.slug}">
          <span class="tag">B</span>
          <h2 class="name">Plano Ventas (Editor)</h2>
          <p class="desc">Editor visual completo de PDF, tabla y estado de unidades.</p>
        </a>`}
        ${onlyInteractive ? '' : `<a class="card" href="/gerente-ventas/${dev.slug}/plano-descargar?dev=${dev.slug}">
          <span class="tag">C</span>
          <h2 class="name">Plano Descargar</h2>
          <p class="desc">Descarga rápida del PDF con carga automática de datos.</p>
        </a>`}
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

  try {
    let pdfBuffer;
    try {
      const browser = await getSharedPdfBrowser();
      pdfBuffer = await buildPdfBufferWithBrowser(browser, html);
    } catch (firstErr) {
      if (!isRetryablePdfError(firstErr)) throw firstErr;
      // If Chromium target dies, recreate browser and retry once.
      try {
        if (sharedPdfBrowser) await sharedPdfBrowser.close();
      } catch {}
      sharedPdfBrowser = null;
      const retryBrowser = await getSharedPdfBrowser();
      pdfBuffer = await buildPdfBufferWithBrowser(retryBrowser, html);
    }

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
  }
});

app.get('/form', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'form.html'));
});

app.get('/form-nacional', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'form-nacional.html'));
});

app.get('/form-nacional-moral', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'form-nacional-moral.html'));
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

app.get('/api/plds/formats-nacional', (req, res) => {
  const items = Object.entries(nationalFormats).map(([id, f]) => ({
    id,
    name: f.name,
    template: f.file,
    htmlRoute: `/format-nacional/${id}`,
    pdfGetRoute: `/format-nacional/${id}/pdf`,
    pdfPostRoute: `/format-nacional/${id}/pdf`
  }));
  res.json({ items });
});

app.get('/api/plds/formats-nacional-moral', (req, res) => {
  const items = Object.entries(nationalMoralFormats).map(([id, f]) => ({
    id,
    name: f.name,
    template: f.file,
    htmlRoute: `/format-nacional-moral/${id}`,
    pdfGetRoute: `/format-nacional-moral/${id}/pdf`,
    pdfPostRoute: `/format-nacional-moral/${id}/pdf`
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

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const data = fs.existsSync(DATA_PATH) ? JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) : {};
  const html = renderTemplate(format.file, data, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format/${id}/pdf`,
    bodyClass: 'print'
  });

  try {
    const pdf = await generatePdfWithSharedBrowser(html, {
      format: 'Letter',
      printBackground: true,
      margin: { top: '0in', right: '0in', bottom: '0in', left: '0in' }
    });

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename=${format.file.replace('.html', '')}.pdf`
    });
    res.send(pdf);
  } catch (err) {
    log(`Error en GET /format/${id}/pdf: ${err && err.stack ? err.stack : err}`);
    res.status(500).send('No se pudo generar el PDF');
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

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const html = renderTemplate(format.file, data, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format/${id}/pdf`,
    bodyClass: 'print'
  });

  try {
    const pdfOpts = {
      format: 'letter',
      printBackground: true,
      margin: { top: '0in', right: '0in', bottom: '0.6in', left: '0in' }
    };
    // FR-VEN-19 footer is rendered in HTML to match FR-VEN-10 layout
    const pdf = await generatePdfWithSharedBrowser(html, pdfOpts);

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename=${format.file.replace('.html', '')}.pdf`
    });
    res.send(pdf);
  } catch (err) {
    log(`Error en POST /format/${id}/pdf: ${err && err.stack ? err.stack : err}`);
    res.status(500).send('No se pudo generar el PDF');
  }
});

app.get('/format-nacional/:id', (req, res) => {
  const id = req.params.id;
  const format = nationalFormats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const data = fs.existsSync(DATA_PATH) ? JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) : {};
  const renderData = withPldsFooterCode(data, 'N', id);
  const html = renderTemplate(format.file, renderData, {
    baseUrl,
    showControls: !req.query.print,
    downloadUrl: `/format-nacional/${id}/pdf`,
    bodyClass: 'interactive'
  });
  res.send(html);
});

app.get('/format-nacional/:id/pdf', async (req, res) => {
  const id = req.params.id;
  const format = nationalFormats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const data = fs.existsSync(DATA_PATH) ? JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) : {};
  const renderData = withPldsFooterCode(data, 'N', id);
  const html = renderTemplate(format.file, renderData, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format-nacional/${id}/pdf`,
    bodyClass: 'print'
  });

  try {
    const pdf = await generatePdfWithSharedBrowser(html, {
      format: 'Letter',
      printBackground: true,
      margin: { top: '0in', right: '0in', bottom: '0in', left: '0in' }
    });

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename=${format.file.replace('.html', '')}.pdf`
    });
    res.send(pdf);
  } catch (err) {
    log(`Error en GET /format-nacional/${id}/pdf: ${err && err.stack ? err.stack : err}`);
    res.status(500).send('No se pudo generar el PDF');
  }
});

app.post('/format-nacional/:id/pdf', async (req, res) => {
  const id = req.params.id;
  const format = nationalFormats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const data = { ...(req.body || {}) };
  if (id === '10' || id === '19' || id === '21' || id === '35') {
    Object.keys(data).forEach((k) => {
      const v = data[k];
      if (typeof v === 'string' && v.trim() === '') data[k] = 'N/A';
    });
  }

  const saved = persistSubmission(`N-${id}`, format.name, data);
  log(`Submission guardado: ${saved.id} formato=N-${id}`);

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const renderData = withPldsFooterCode(data, 'N', id);
  const html = renderTemplate(format.file, renderData, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format-nacional/${id}/pdf`,
    bodyClass: 'print'
  });

  try {
    const pdfOpts = {
      format: 'letter',
      printBackground: true,
      margin: { top: '0in', right: '0in', bottom: '0.6in', left: '0in' }
    };
    const pdf = await generatePdfWithSharedBrowser(html, pdfOpts);

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename=${format.file.replace('.html', '')}.pdf`
    });
    res.send(pdf);
  } catch (err) {
    log(`Error en POST /format-nacional/${id}/pdf: ${err && err.stack ? err.stack : err}`);
    res.status(500).send('No se pudo generar el PDF');
  }
});

app.get('/format-nacional-moral/:id', (req, res) => {
  const id = req.params.id;
  const format = nationalMoralFormats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const data = fs.existsSync(DATA_PATH) ? JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) : {};
  const renderData = withPldsFooterCode(data, 'NM', id);
  const html = renderTemplate(format.file, renderData, {
    baseUrl,
    showControls: !req.query.print,
    downloadUrl: `/format-nacional-moral/${id}/pdf`,
    bodyClass: 'interactive'
  });
  res.send(html);
});

app.get('/format-nacional-moral/:id/pdf', async (req, res) => {
  const id = req.params.id;
  const format = nationalMoralFormats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const data = fs.existsSync(DATA_PATH) ? JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) : {};
  const renderData = withPldsFooterCode(data, 'NM', id);
  const html = renderTemplate(format.file, renderData, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format-nacional-moral/${id}/pdf`,
    bodyClass: 'print'
  });

  try {
    const pdf = await generatePdfWithSharedBrowser(html, {
      format: 'Letter',
      printBackground: true,
      margin: { top: '0in', right: '0in', bottom: '0in', left: '0in' }
    });

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename=${format.file.replace('.html', '')}.pdf`
    });
    res.send(pdf);
  } catch (err) {
    log(`Error en GET /format-nacional-moral/${id}/pdf: ${err && err.stack ? err.stack : err}`);
    res.status(500).send('No se pudo generar el PDF');
  }
});

app.post('/format-nacional-moral/:id/pdf', async (req, res) => {
  const id = req.params.id;
  const format = nationalMoralFormats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const data = { ...(req.body || {}) };
  if (id === '10' || id === '19' || id === '21' || id === '35') {
    Object.keys(data).forEach((k) => {
      const v = data[k];
      if (typeof v === 'string' && v.trim() === '') data[k] = 'N/A';
    });
  }

  const saved = persistSubmission(`NM-${id}`, format.name, data);
  log(`Submission guardado: ${saved.id} formato=NM-${id}`);

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const renderData = withPldsFooterCode(data, 'NM', id);
  const html = renderTemplate(format.file, renderData, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format-nacional-moral/${id}/pdf`,
    bodyClass: 'print'
  });

  try {
    const pdfOpts = {
      format: 'letter',
      printBackground: true,
      margin: { top: '0in', right: '0in', bottom: '0.6in', left: '0in' }
    };
    const pdf = await generatePdfWithSharedBrowser(html, pdfOpts);

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename=${format.file.replace('.html', '')}.pdf`
    });
    res.send(pdf);
  } catch (err) {
    log(`Error en POST /format-nacional-moral/${id}/pdf: ${err && err.stack ? err.stack : err}`);
    res.status(500).send('No se pudo generar el PDF');
  }
});

async function startServer() {
  try {
    await ensureWhisperlistStorageReady();
  } catch (err) {
    log(`Fallo inicializando Whisperlist storage: ${err && err.stack ? err.stack : err}`);
    process.exit(1);
  }
  const server = app.listen(PORT, HOST, () => {
    log(`Servidor listo en http://${HOST}:${PORT}`);
  });
  server.on('error', (err) => {
    log(`Error al iniciar servidor: ${err && err.stack ? err.stack : err}`);
  });
}

startServer();
