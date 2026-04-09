const express = require('express');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execSync } = require('child_process');
process.env.PUPPETEER_CACHE_DIR = path.join(__dirname, '.cache', 'puppeteer');
const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const { PDFDocument } = require('pdf-lib');
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
EXTRA_ALLOWED_EMAILS.add('jmotta@relatedgroup.com');
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
function isWritableDir(dirPath) {
  try {
    if (!dirPath) return false;
    fs.mkdirSync(dirPath, { recursive: true });
    const probe = path.join(dirPath, `.write-test-${process.pid}-${Date.now()}`);
    fs.writeFileSync(probe, 'ok', 'utf-8');
    fs.unlinkSync(probe);
    return true;
  } catch {
    return false;
  }
}

function resolveDataDir() {
  const envDir = String(process.env.APP_DATA_DIR || '').trim();
  if (envDir) return path.resolve(envDir);
  const candidates = [
    '/var/data',
    path.join(os.tmpdir(), 'pldssimca-data'),
    REPO_DATA_DIR
  ];
  for (const candidate of candidates) {
    if (isWritableDir(candidate)) return path.resolve(candidate);
  }
  return path.resolve(REPO_DATA_DIR);
}

const DATA_DIR = resolveDataDir();
const DATA_PATH = path.join(DATA_DIR, 'sample.json');
const ROI_MASTER_CSV_PATH = path.join(DATA_DIR, 'roi-master.csv');
const SOLAR_MIDTOWN_BROCHURE_ES_ENV_PATH = String(process.env.SOLAR_MIDTOWN_BROCHURE_ES_PATH || '').trim();
const SOLAR_MIDTOWN_BROCHURE_EN_ENV_PATH = String(process.env.SOLAR_MIDTOWN_BROCHURE_EN_PATH || '').trim();
const SOLAR_MIDTOWN_BROCHURE_ES_ENV_URL = String(process.env.SOLAR_MIDTOWN_BROCHURE_ES_URL || '').trim();
const SOLAR_MIDTOWN_BROCHURE_EN_ENV_URL = String(process.env.SOLAR_MIDTOWN_BROCHURE_EN_URL || '').trim();
const SOLAR_MIDTOWN_BROCHURE_PATH = path.join(os.homedir(), 'Downloads', 'SOLAR Midtown ESP.pdf');
const SOLAR_MIDTOWN_BROCHURE_ENG_PATH = path.join(os.homedir(), 'Downloads', 'SOLAR Midtown ENG.pdf');
const SOLAR_MIDTOWN_BROCHURE_ES_FALLBACK_URL = 'https://raw.githubusercontent.com/ellocodivision/pldssimca/main/data/SOLAR%20Midtown%20ESP.pdf';
const SOLAR_MIDTOWN_BROCHURE_EN_FALLBACK_URL = 'https://raw.githubusercontent.com/ellocodivision/pldssimca/main/data/SOLAR%20Midtown%20ENG.pdf';
const SOLAR_MIDTOWN_LAYOUT_PATH = path.join(DATA_DIR, 'presentaciones-solar-midtown-layout.json');
const SOLAR_MIDTOWN_CROPS_PATH = path.join(DATA_DIR, 'presentaciones-solar-midtown-crops.json');
const SOLAR_MIDTOWN_LAYOUT_REPO_PATH = path.join(REPO_DATA_DIR, 'presentaciones-solar-midtown-layout.json');
const SOLAR_MIDTOWN_CROPS_REPO_PATH = path.join(REPO_DATA_DIR, 'presentaciones-solar-midtown-crops.json');
const SOLAR_MIDTOWN_LAYOUT_SEED_PATH = path.join(__dirname, 'seed-data', 'presentaciones-solar-midtown-layout.json');
const SOLAR_MIDTOWN_CROPS_SEED_PATH = path.join(__dirname, 'seed-data', 'presentaciones-solar-midtown-crops.json');
const TABLA_PAGOS_LAYOUT_PATH = path.join(DATA_DIR, 'tabla-pagos-layout.json');
const VICEROY_PAYMENT_PLAN_LAYOUT_PATH = path.join(DATA_DIR, 'viceroy-payment-plan-layout.json');
const BROKERS_SIMCA_TEMPLATE_PATH = path.join(DATA_DIR, 'brokers-simca-template.json');
const CEIBA_CROPS_PATH = path.join(DATA_DIR, 'presentaciones-ceiba-crops.json');
const CEIBA_CROPS_REPO_PATH = path.join(REPO_DATA_DIR, 'presentaciones-ceiba-crops.json');
const SOLAR_MIDTOWN_APPENDIX_CHUNK_SIZE = clampNumber(Number(process.env.SOLAR_MIDTOWN_APPENDIX_CHUNK_SIZE), 20, 5, 60);
const SOLAR_MIDTOWN_PLAN_FETCH_CONCURRENCY = clampNumber(Number(process.env.SOLAR_MIDTOWN_PLAN_FETCH_CONCURRENCY), 6, 1, 16);
const SOLAR_MIDTOWN_EDITOR_EMAIL = String(process.env.SOLAR_MIDTOWN_EDITOR_EMAIL || 'martin@simca.mx').toLowerCase();
const CEIBA_EDITOR_EMAIL = String(process.env.CEIBA_EDITOR_EMAIL || SOLAR_MIDTOWN_EDITOR_EMAIL || 'martin@simca.mx').toLowerCase();
const SUBMISSIONS_PATH = path.join(DATA_DIR, 'submissions.json');
const OWNER_SERVICES_PATH = path.join(DATA_DIR, 'owner-services.json');
const WHISPERLIST_JSON_PATH = path.join(DATA_DIR, 'viceroy-whisperlist.json');
const WHISPERLIST_EXCEL_PATH = path.join(DATA_DIR, 'VICEROY WHISPERLIST.xlsx');
const VICEROY_REGISTROS_JSON_PATH = path.join(DATA_DIR, 'viceroy-registros.json');
const VICEROY_REGISTROS_EXCEL_PATH = path.join(DATA_DIR, 'VICEROY REGISTROS.xlsx');
const VICEROY_ROOM_RESERVATIONS_PATH = path.join(DATA_DIR, 'viceroy-room-reservations.json');
const VICEROY_PILOTO_CONFIG_NAME = 'viceroy-tipologias.json';
const FLOOR_JSON_DIR = path.join(DATA_DIR, 'plano-ventas-floors');
const DEVELOPMENTS_DIR = path.join(DATA_DIR, 'developments');
const SEED_DEVELOPMENTS_DIR = path.join(__dirname, 'seed-data', 'developments');
const DEFAULT_DEVELOPMENT_SLUG = 'ceiba';
const FLOOR_JSON_FILE_RE = /^unidades-marcadas(?:\s*\(\d+\))?\.json$/i;
const FLOOR_MAPPED_JSON_FILE_RE = /^imagen-mapeada-.*\.json$/i;
log(`DATA_DIR activo: ${DATA_DIR}`);

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
  '26': { name: 'FR-VEN-26 Origen de los Recursos (Nacional PF)', file: 'format-26-nacional-pf.html' },
  '19': { name: 'FR-VEN-19 Beneficiario Controlador PF (Nacional)', file: 'format-19-nacional.html' },
  '10': { name: 'FR-VEN-10 Identificación del Cliente (Nacional PF)', file: 'format-10-nacional.html' }
};
const nationalMoralFormats = {
  '35': { name: 'FR-VEN-35 Aviso de Privacidad (Nacional PM)', file: 'format-35-nacional.html' },
  '21': { name: 'FR-VEN-21 Consulta de Listas de Personas Bloqueadas', file: 'format-21.html' },
  '26': { name: 'FR-VEN-26 Origen de los Recursos (Nacional PM)', file: 'format-26-nacional-pm.html' },
  '19': { name: 'FR-VEN-19 Beneficiario Controlador PM (Nacional)', file: 'format-19-nacional-moral.html' },
  '10': { name: 'FR-VEN-10 Identificación del Cliente (Nacional PM)', file: 'format-10-nacional-moral.html' }
};
const foreignMoralFormats = {
  '35': { name: 'FR-VEN-35 Aviso de Privacidad (Extranjera PM)', file: 'format-35.html' },
  '21': { name: 'FR-VEN-21 Consulta de Listas de Personas Bloqueadas', file: 'format-21.html' },
  '26': { name: 'FR-VEN-26 Origen de los Recursos (Extranjera PM)', file: 'format-26-extranjera-moral.html' },
  '19': { name: 'FR-VEN-19 Beneficiario Controlador PM (Extranjera)', file: 'format-19-extranjera-moral.html' },
  '10': { name: 'FR-VEN-10 Identificación del Cliente (Extranjera PM)', file: 'format-10-extranjera-moral.html' }
};
const WHISPERLIST_CANALES = ['SIMCA', 'RELATED'];
const WHISPERLIST_TIPOS_VENTA = ['EXTERNO', 'INTERNO', 'MERITO PROPIO'];
const WHISPERLIST_RECAMARAS = ['1', '2', '3', '2PH', '3PH'];
const WHISPERLIST_PAISES = ['MEXICO', 'CANADA', 'USA', 'RESTO AMERICA', 'EUROPA', 'ASIA', 'AFRICA', 'OCEANIA'];

const whisperlistPool = USE_WHISPERLIST_DB
  ? new Pool({
    connectionString: DATABASE_URL,
    ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false
  })
  : null;
let whisperlistStorageReady = false;
let viceroyRegistrosStorageReady = false;

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
  return res.redirect('/viceroy');
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

function hasSolarMidtownEditorAccess(req) {
  if (LOCAL_NO_AUTH) return true;
  const email = String(req && req.user && req.user.email || '').trim().toLowerCase();
  return Boolean(email && email === SOLAR_MIDTOWN_EDITOR_EMAIL);
}

function requireSolarMidtownEditor(req, res, next) {
  if (hasSolarMidtownEditorAccess(req)) return next();
  if (String(req.path || '').startsWith('/api/')) {
    return res.status(403).json({ error: 'Solo la cuenta autorizada puede editar plantilla/cortes.' });
  }
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
      <h1>Acceso restringido</h1>
      <p>Solo la cuenta autorizada puede editar plantilla y cortes de Solar Midtown.</p>
      <p>Cuenta autorizada: <strong>${escapeHtml(SOLAR_MIDTOWN_EDITOR_EMAIL)}</strong>.</p>
    </section>
  </body></html>`);
}

function hasCeibaEditorAccess(req) {
  if (LOCAL_NO_AUTH) return true;
  const email = String(req && req.user && req.user.email || '').trim().toLowerCase();
  return Boolean(email && email === CEIBA_EDITOR_EMAIL);
}

function requireCeibaEditor(req, res, next) {
  if (hasCeibaEditorAccess(req)) return next();
  if (String(req.path || '').startsWith('/api/')) {
    return res.status(403).json({ error: 'Solo la cuenta autorizada puede editar cortes de Ceiba.' });
  }
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
      <h1>Acceso restringido</h1>
      <p>Solo la cuenta autorizada puede editar cortes de Ceiba.</p>
      <p>Cuenta autorizada: <strong>${escapeHtml(CEIBA_EDITOR_EMAIL)}</strong>.</p>
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
  const copyAlwaysIfExists = (src, dst) => {
    try {
      if (!fs.existsSync(src)) return;
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
    const repoFloorDir = path.join(REPO_DATA_DIR, 'developments', dev.slug, 'plano-ventas-floors');
    try {
      fs.readdirSync(repoFloorDir).forEach((name) => {
        const lower = String(name || '').toLowerCase();
        if (!lower.endsWith('.json')) return;
        copyIfMissing(
          path.join(repoFloorDir, name),
          path.join(floorDir, name)
        );
      });
    } catch {}
  });
  if (!fs.existsSync(SUBMISSIONS_PATH)) fs.writeFileSync(SUBMISSIONS_PATH, '[]', 'utf-8');
  if (!fs.existsSync(DATA_PATH)) fs.writeFileSync(DATA_PATH, '{}', 'utf-8');
  if (!fs.existsSync(VICEROY_ROOM_RESERVATIONS_PATH)) {
    writeJson(VICEROY_ROOM_RESERVATIONS_PATH, { rows: [], updatedAt: null });
  }
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

  // Keep Solar Midtown presets identical to repo on every boot (local/dev/prod).
  copyAlwaysIfExists(
    SOLAR_MIDTOWN_LAYOUT_REPO_PATH,
    SOLAR_MIDTOWN_LAYOUT_PATH
  );
  copyAlwaysIfExists(
    SOLAR_MIDTOWN_LAYOUT_SEED_PATH,
    SOLAR_MIDTOWN_LAYOUT_PATH
  );
  copyAlwaysIfExists(
    SOLAR_MIDTOWN_CROPS_REPO_PATH,
    SOLAR_MIDTOWN_CROPS_PATH
  );
  copyAlwaysIfExists(
    SOLAR_MIDTOWN_CROPS_SEED_PATH,
    SOLAR_MIDTOWN_CROPS_PATH
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
  const repoFloorDir = path.join(REPO_DATA_DIR, 'developments', devSlug, 'plano-ventas-floors');
  if (devSlug === DEFAULT_DEVELOPMENT_SLUG) {
    return [primary, FLOOR_JSON_DIR];
  }
  if (devSlug === 'viceroy-piloto') {
    const seedFloorDir = path.join(SEED_DEVELOPMENTS_DIR, devSlug, 'plano-ventas-floors');
    return [primary, repoFloorDir, FLOOR_JSON_DIR, seedFloorDir];
  }
  return [primary, repoFloorDir];
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
    selectedFloorJsonName: '',
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
  const selectedFloorJsonName = String(source.selectedFloorJsonName || '').trim()
    ? sanitizeJsonFileName(source.selectedFloorJsonName)
    : '';
  return {
    tipologias,
    unitTipologiaMap,
    unitRecamarasMap,
    mapFloorOrder,
    selectedFloorJsonName,
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

function updateViceroyPilotoSelectedFloorJson(fileName) {
  const current = readViceroyPilotoConfig();
  const next = {
    ...current,
    selectedFloorJsonName: String(fileName || '').trim()
      ? sanitizeJsonFileName(fileName)
      : ''
  };
  return writeViceroyPilotoConfig(next);
}

function listFloorJsonFiles(dir) {
  try {
    return fs.readdirSync(dir)
      .filter((name) => {
        const lower = String(name || '').toLowerCase();
        if (!lower.endsWith('.json')) return false;
        if (/^version-.*\.json$/i.test(name)) return true;
        if (FLOOR_MAPPED_JSON_FILE_RE.test(name)) return true;
        return FLOOR_JSON_FILE_RE.test(name);
      })
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
      '--disable-gpu',
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
    // Avoid hanging on long-lived/in-flight asset requests; readiness is handled via __pdfReady.
    await page.setContent(html, { waitUntil: 'domcontentloaded', timeout: 45000 });
    await page.emulateMediaType('print');
    try {
      const hasReadyMarker = await page.evaluate(() => Object.prototype.hasOwnProperty.call(window, '__pdfReady'));
      if (hasReadyMarker) {
        await page.waitForFunction(() => window.__pdfReady === true, { timeout: 90000 });
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

function csvCellToString(raw) {
  if (raw == null) return '';
  if (typeof raw === 'number' && Number.isFinite(raw)) return String(raw);
  return String(raw).trim();
}

function parseCurrencyLike(raw) {
  if (raw == null || raw === '') return null;
  if (typeof raw === 'number' && Number.isFinite(raw)) return raw;
  const cleaned = String(raw)
    .replace(/[$,%\s]/g, '')
    .replace(/,/g, '');
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : null;
}

function parsePercentLike(raw) {
  if (raw == null || raw === '') return null;
  if (typeof raw === 'number' && Number.isFinite(raw)) {
    return raw <= 1 ? raw * 100 : raw;
  }
  const text = String(raw).trim();
  const hasPercent = text.includes('%');
  const base = parseCurrencyLike(text);
  if (base == null) return null;
  if (hasPercent) return base;
  return base <= 1 ? base * 100 : base;
}

function formatCurrency(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return '';
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 2
  }).format(num);
}

function normalizeDevLabel(raw) {
  return String(raw || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function makeSolarMidtownRowId(unidad, rowNumber) {
  const unitKey = normalizeDevLabel(unidad).replace(/[^A-Z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  if (unitKey) return `solar-mt-unit-${unitKey}`;
  return `solar-mt-${Number(rowNumber) || 0}`;
}

function makeCeibaRowId(unidad, rowNumber) {
  const unitKey = normalizeDevLabel(unidad).replace(/[^A-Z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  if (unitKey) return `ceiba-unit-${unitKey}`;
  return `ceiba-${Number(rowNumber) || 0}`;
}

function readSolarMidtownRowsFromCbs() {
  if (!fs.existsSync(ROI_MASTER_CSV_PATH)) {
    return {
      rows: [],
      sourceFile: ROI_MASTER_CSV_PATH,
      error: 'No se encontró CSV maestro (roi-master.csv).'
    };
  }
  let workbook;
  try {
    workbook = XLSX.readFile(ROI_MASTER_CSV_PATH, { cellDates: true });
  } catch (err) {
    return {
      rows: [],
      sourceFile: ROI_MASTER_CSV_PATH,
      error: `No se pudo leer CSV maestro: ${err && err.message ? err.message : 'error desconocido'}`
    };
  }
  const firstSheetName = workbook.SheetNames && workbook.SheetNames[0];
  if (!firstSheetName || !workbook.Sheets[firstSheetName]) {
    return {
      rows: [],
      sourceFile: ROI_MASTER_CSV_PATH,
      error: 'No hay hoja utilizable en el CSV maestro.'
    };
  }
  const sheet = workbook.Sheets[firstSheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
  const out = [];
  const cropMap = readSolarMidtownCropMap();
  matrix.forEach((row, idx) => {
    if (!Array.isArray(row)) return;
    if (idx === 0) return;
    const desarrollo = csvCellToString(row[2]);
    if (normalizeDevLabel(desarrollo) !== 'SOLAR MT') return;

    const planLink = csvCellToString(row[1]);
    const unidad = csvCellToString(row[3]) || csvCellToString(row[4]);
    const colE = csvCellToString(row[4]);
    const colF = csvCellToString(row[5]);
    const colG = csvCellToString(row[6]);

    const precioListaRaw = row[7];
    const descuentoPctRaw = row[8];
    const descuentoValorRaw = row[9];
    const precioFinalRaw = row[10];

    const precioLista = parseCurrencyLike(precioListaRaw);
    let descuentoPct = parsePercentLike(descuentoPctRaw);
    let descuentoValor = parseCurrencyLike(descuentoValorRaw);
    let precioFinal = parseCurrencyLike(precioFinalRaw);

    if (precioFinal == null && precioLista != null && descuentoValor != null) {
      precioFinal = precioLista - descuentoValor;
    }
    if (precioFinal == null && precioLista != null && descuentoPct != null) {
      precioFinal = precioLista * (1 - (descuentoPct / 100));
    }
    if (descuentoValor == null && precioLista != null && precioFinal != null) {
      descuentoValor = precioLista - precioFinal;
    }
    if (descuentoPct == null && precioLista != null && descuentoValor != null && precioLista !== 0) {
      descuentoPct = (descuentoValor / precioLista) * 100;
    }

    const finalForPayments = Number.isFinite(precioFinal)
      ? precioFinal
      : (Number.isFinite(precioLista) ? precioLista : 0);
    const enganche30 = finalForPayments * 0.30;
    const seisMeses1 = finalForPayments * 0.10;
    const seisMeses2 = finalForPayments * 0.10;
    const entrega50 = finalForPayments * 0.50;

    const rowId = makeSolarMidtownRowId(unidad, idx + 1);
    out.push({
      id: rowId,
      rowNumber: idx + 1,
      planLink,
      desarrollo,
      unidad,
      crop: getSolarMidtownCropForId(rowId, cropMap),
      colE,
      colF,
      colG,
      precioLista,
      descuentoPct,
      descuentoValor,
      precioFinal,
      precioListaFmt: formatCurrency(precioLista),
      descuentoPctFmt: Number.isFinite(descuentoPct) ? `${descuentoPct.toFixed(2)}%` : '',
      descuentoValorFmt: formatCurrency(descuentoValor),
      precioFinalFmt: formatCurrency(precioFinal),
      pagos: {
        enganche30,
        seisMeses1,
        seisMeses2,
        entrega50,
        enganche30Fmt: formatCurrency(enganche30),
        seisMeses1Fmt: formatCurrency(seisMeses1),
        seisMeses2Fmt: formatCurrency(seisMeses2),
        entrega50Fmt: formatCurrency(entrega50)
      }
    });
  });
  return {
    rows: out,
    sourceFile: ROI_MASTER_CSV_PATH,
    error: ''
  };
}

function readCeibaRowsFromCbs() {
  if (!fs.existsSync(ROI_MASTER_CSV_PATH)) {
    return {
      rows: [],
      sourceFile: ROI_MASTER_CSV_PATH,
      error: 'No se encontró CSV maestro (roi-master.csv).'
    };
  }
  let workbook;
  try {
    workbook = XLSX.readFile(ROI_MASTER_CSV_PATH, { cellDates: true });
  } catch (err) {
    return {
      rows: [],
      sourceFile: ROI_MASTER_CSV_PATH,
      error: `No se pudo leer CSV maestro: ${err && err.message ? err.message : 'error desconocido'}`
    };
  }
  const firstSheetName = workbook.SheetNames && workbook.SheetNames[0];
  if (!firstSheetName || !workbook.Sheets[firstSheetName]) {
    return {
      rows: [],
      sourceFile: ROI_MASTER_CSV_PATH,
      error: 'No hay hoja utilizable en el CSV maestro.'
    };
  }
  const sheet = workbook.Sheets[firstSheetName];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
  const out = [];
  const cropMap = readCeibaCropMap();
  matrix.forEach((row, idx) => {
    if (!Array.isArray(row) || idx === 0) return;
    const desarrollo = csvCellToString(row[2]);
    if (normalizeDevLabel(desarrollo) !== 'CEIBA') return;

    const planLink = csvCellToString(row[1]);
    const unidad = csvCellToString(row[3]) || csvCellToString(row[4]);
    const colE = csvCellToString(row[4]);
    const colF = csvCellToString(row[5]);
    const colG = csvCellToString(row[6]);

    const precioListaRaw = row[7];
    const descuentoPctRaw = row[8];
    const descuentoValorRaw = row[9];
    const precioFinalRaw = row[10];

    const precioLista = parseCurrencyLike(precioListaRaw);
    let descuentoPct = parsePercentLike(descuentoPctRaw);
    let descuentoValor = parseCurrencyLike(descuentoValorRaw);
    let precioFinal = parseCurrencyLike(precioFinalRaw);

    if (precioFinal == null && precioLista != null && descuentoValor != null) {
      precioFinal = precioLista - descuentoValor;
    }
    if (precioFinal == null && precioLista != null && descuentoPct != null) {
      precioFinal = precioLista * (1 - (descuentoPct / 100));
    }
    if (descuentoValor == null && precioLista != null && precioFinal != null) {
      descuentoValor = precioLista - precioFinal;
    }
    if (descuentoPct == null && precioLista != null && descuentoValor != null && precioLista !== 0) {
      descuentoPct = (descuentoValor / precioLista) * 100;
    }

    const rowId = makeCeibaRowId(unidad, idx + 1);
    out.push({
      id: rowId,
      rowNumber: idx + 1,
      planLink,
      desarrollo,
      unidad,
      crop: getCeibaCropForId(rowId, cropMap),
      colE,
      colF,
      colG,
      precioLista,
      descuentoPct,
      descuentoValor,
      precioFinal,
      precioListaFmt: formatCurrency(precioLista),
      descuentoPctFmt: Number.isFinite(descuentoPct) ? `${descuentoPct.toFixed(2)}%` : '',
      descuentoValorFmt: formatCurrency(descuentoValor),
      precioFinalFmt: formatCurrency(precioFinal)
    });
  });

  out.sort((a, b) => String(a.unidad || '').localeCompare(String(b.unidad || ''), 'es', { numeric: true, sensitivity: 'base' }));
  return { rows: out, sourceFile: ROI_MASTER_CSV_PATH, error: '' };
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

function sha256ForString(value) {
  return crypto.createHash('sha256').update(String(value || ''), 'utf8').digest('hex');
}

function sha256ForJson(value) {
  return sha256ForString(JSON.stringify(value == null ? null : value));
}

function sha256ForFile(filePath) {
  try {
    if (!fs.existsSync(filePath)) return '';
    return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
  } catch {
    return '';
  }
}

function readFirstExistingJson(paths, fallback) {
  const list = Array.isArray(paths) ? paths : [];
  const unique = Array.from(new Set(list.map((p) => String(p || '').trim()).filter(Boolean)));
  for (const p of unique) {
    try {
      if (!fs.existsSync(p)) continue;
      return readJson(p, fallback);
    } catch {}
  }
  return fallback;
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function deepCloneJson(value) {
  return JSON.parse(JSON.stringify(value));
}

function isJsonEqual(a, b) {
  return JSON.stringify(a) === JSON.stringify(b);
}

function applyNonDefaultOverlay(baseInput, overlayInput, defaultsInput) {
  const base = isPlainObject(baseInput) ? baseInput : {};
  const overlay = isPlainObject(overlayInput) ? overlayInput : {};
  const defaults = isPlainObject(defaultsInput) ? defaultsInput : {};
  Object.keys(overlay).forEach((key) => {
    const overlayVal = overlay[key];
    const defaultVal = defaults[key];
    const baseVal = base[key];
    if (isPlainObject(overlayVal) && isPlainObject(defaultVal)) {
      const nextBase = isPlainObject(baseVal) ? baseVal : {};
      base[key] = applyNonDefaultOverlay(nextBase, overlayVal, defaultVal);
      return;
    }
    if (!isJsonEqual(overlayVal, defaultVal)) {
      base[key] = overlayVal;
    }
  });
  return base;
}

function defaultSolarMidtownPlanCrop() {
  return { x: 0, y: 0, w: 100, h: 100 };
}

function defaultCeibaPlanCrop() {
  return { x: 0, y: 0, w: 100, h: 100 };
}

function normalizeSolarMidtownPlanCrop(input) {
  const base = defaultSolarMidtownPlanCrop();
  const src = input && typeof input === 'object' ? input : {};
  // Backward compatibility with previous format {x,y,zoom}
  if (Object.prototype.hasOwnProperty.call(src, 'zoom') && !Object.prototype.hasOwnProperty.call(src, 'w')) {
    const zoom = clampNumber(src.zoom, 1, 1, 4);
    const boxW = clampNumber(100 / zoom, 100, 5, 100);
    const boxH = clampNumber(100 / zoom, 100, 5, 100);
    const centerX = clampNumber(src.x, 50, -300, 300);
    const centerY = clampNumber(src.y, 50, -300, 300);
    const left = centerX - (boxW / 2);
    const top = centerY - (boxH / 2);
    return { x: left, y: top, w: boxW, h: boxH };
  }
  const width = clampNumber(src.w, base.w, 5, 100);
  const height = clampNumber(src.h, base.h, 5, 100);
  return {
    x: clampNumber(src.x, base.x, -300, 300),
    y: clampNumber(src.y, base.y, -300, 300),
    w: width,
    h: height
  };
}

function normalizeCeibaPlanCrop(input) {
  const base = defaultCeibaPlanCrop();
  const src = input && typeof input === 'object' ? input : {};
  if (Object.prototype.hasOwnProperty.call(src, 'zoom') && !Object.prototype.hasOwnProperty.call(src, 'w')) {
    const zoom = clampNumber(src.zoom, 1, 1, 4);
    const boxW = clampNumber(100 / zoom, 100, 5, 100);
    const boxH = clampNumber(100 / zoom, 100, 5, 100);
    const centerX = clampNumber(src.x, 50, -300, 300);
    const centerY = clampNumber(src.y, 50, -300, 300);
    const left = centerX - (boxW / 2);
    const top = centerY - (boxH / 2);
    return { x: left, y: top, w: boxW, h: boxH };
  }
  const width = clampNumber(src.w, base.w, 5, 100);
  const height = clampNumber(src.h, base.h, 5, 100);
  return {
    x: clampNumber(src.x, base.x, -300, 300),
    y: clampNumber(src.y, base.y, -300, 300),
    w: width,
    h: height
  };
}

function normalizeSolarMidtownCropMap(input) {
  const src = input && typeof input === 'object' ? input : {};
  const out = {};
  Object.keys(src).forEach((key) => {
    const id = String(key || '').trim();
    if (!id) return;
    out[id] = normalizeSolarMidtownPlanCrop(src[id]);
  });
  return out;
}

function normalizeCeibaCropMap(input) {
  const src = input && typeof input === 'object' ? input : {};
  const out = {};
  Object.keys(src).forEach((key) => {
    const id = String(key || '').trim();
    if (!id) return;
    out[id] = normalizeCeibaPlanCrop(src[id]);
  });
  return out;
}

function isDefaultSolarMidtownCrop(cropInput) {
  const crop = normalizeSolarMidtownPlanCrop(cropInput);
  const def = defaultSolarMidtownPlanCrop();
  return crop.x === def.x && crop.y === def.y && crop.w === def.w && crop.h === def.h;
}

function isDefaultCeibaCrop(cropInput) {
  const crop = normalizeCeibaPlanCrop(cropInput);
  const def = defaultCeibaPlanCrop();
  return crop.x === def.x && crop.y === def.y && crop.w === def.w && crop.h === def.h;
}

function readSolarMidtownCropMap() {
  const repoRaw = readJson(SOLAR_MIDTOWN_CROPS_REPO_PATH, {});
  const seedRaw = readJson(SOLAR_MIDTOWN_CROPS_SEED_PATH, {});
  const runtimeRaw = readJson(SOLAR_MIDTOWN_CROPS_PATH, {});
  const repoMap = normalizeSolarMidtownCropMap(repoRaw);
  const seedMap = normalizeSolarMidtownCropMap(seedRaw);
  const runtimeMap = normalizeSolarMidtownCropMap(runtimeRaw);
  if (Object.keys(repoMap).length > 0) {
    return repoMap;
  }
  if (Object.keys(seedMap).length > 0) {
    return seedMap;
  }
  const merged = { ...repoMap, ...seedMap };
  Object.keys(runtimeMap).forEach((id) => {
    const runtimeCrop = runtimeMap[id];
    const repoCrop = merged[id];
    if (!repoCrop) {
      merged[id] = runtimeCrop;
      return;
    }
    // If runtime has a real edit, keep it. If runtime is default/no-crop, preserve repo edit.
    if (!isDefaultSolarMidtownCrop(runtimeCrop) || isDefaultSolarMidtownCrop(repoCrop)) {
      merged[id] = runtimeCrop;
    }
  });
  return merged;
}

function readCeibaCropMap() {
  const repoRaw = readJson(CEIBA_CROPS_REPO_PATH, {});
  const runtimeRaw = readJson(CEIBA_CROPS_PATH, {});
  const repoMap = normalizeCeibaCropMap(repoRaw);
  const runtimeMap = normalizeCeibaCropMap(runtimeRaw);
  const merged = { ...repoMap };
  Object.keys(runtimeMap).forEach((id) => {
    const runtimeCrop = runtimeMap[id];
    const repoCrop = merged[id];
    if (!repoCrop) {
      merged[id] = runtimeCrop;
      return;
    }
    if (!isDefaultCeibaCrop(runtimeCrop) || isDefaultCeibaCrop(repoCrop)) {
      merged[id] = runtimeCrop;
    }
  });
  return merged;
}

function saveSolarMidtownCropMap(map) {
  const normalized = normalizeSolarMidtownCropMap(map);
  writeJson(SOLAR_MIDTOWN_CROPS_PATH, normalized);
  return normalized;
}

function saveCeibaCropMap(map) {
  const normalized = normalizeCeibaCropMap(map);
  writeJson(CEIBA_CROPS_PATH, normalized);
  return normalized;
}

function getSolarMidtownCropForId(rowId, cropMap) {
  const id = String(rowId || '').trim();
  if (!id) return defaultSolarMidtownPlanCrop();
  const map = cropMap && typeof cropMap === 'object' ? cropMap : readSolarMidtownCropMap();
  return normalizeSolarMidtownPlanCrop(map[id]);
}

function getCeibaCropForId(rowId, cropMap) {
  const id = String(rowId || '').trim();
  if (!id) return defaultCeibaPlanCrop();
  const map = cropMap && typeof cropMap === 'object' ? cropMap : readCeibaCropMap();
  return normalizeCeibaPlanCrop(map[id]);
}

function buildSolarMidtownCropDataAttrs(cropInput, row, layoutInput) {
  const crop = normalizeSolarMidtownPlanCrop(cropInput);
  const layout = normalizeSolarMidtownLayout(layoutInput);
  const verticalExpandRatio = clampNumber(layout && layout.crop && layout.crop.verticalExpandRatio, 1.9, 1, 2.4);
  const expand = isSolarMidtownVerticalUnit(row && row.unidad) ? verticalExpandRatio : 1;
  return `data-crop-x="${crop.x}" data-crop-y="${crop.y}" data-crop-w="${crop.w}" data-crop-h="${crop.h}" data-crop-expand="${expand}"`;
}

function defaultSolarMidtownLayout() {
  return {
    version: 1,
    crop: {
      verticalExpandRatio: 1.9
    },
    page: {
      backgroundColor: '#f28b2c',
      padding: 16
    },
    title: {
      show: true,
      text: 'Página adicional',
      color: '#1f2330',
      fontSize: 22,
      fontWeight: 700
    },
    subtitle: {
      show: true,
      color: '#5f6572',
      fontSize: 14
    },
    grid: {
      gap: 14,
      leftRatio: 1.2,
      rightRatio: 1
    },
    planBox: {
      backgroundColor: '#f8f8f8',
      borderColor: '#d4d7df',
      borderRadius: 12,
      minHeight: 340
    },
    priceBox: {
      backgroundColor: '#f28b2c',
      borderColor: '#df7d24',
      borderRadius: 12,
      textColor: '#ffffff',
      headingColor: '#ffffff',
      dividerColor: 'rgba(255,255,255,0.35)',
      headingSize: 15,
      bodySize: 12,
      labelSize: 12,
      valueSize: 12,
      metaSize: 12,
      departmentLabelSize: 12,
      departmentValueSize: 12,
      metersLabelSize: 12,
      metersValueSize: 12,
      rowPaddingY: 7,
      lineHeight: 1.2,
      wordSpacing: 0,
      letterSpacing: 0,
      metaTopOffset: 0,
      metaGroupGap: 14,
      verticalAlign: 'start'
    }
  };
}

function clampNumber(value, fallback, min, max) {
  const num = Number(value);
  if (!Number.isFinite(num)) return fallback;
  if (Number.isFinite(min) && num < min) return min;
  if (Number.isFinite(max) && num > max) return max;
  return num;
}

function normalizeSolarMidtownLayout(input) {
  const base = defaultSolarMidtownLayout();
  const src = input && typeof input === 'object' ? input : {};
  const out = {
    ...base,
    crop: {
      ...base.crop,
      ...(src.crop && typeof src.crop === 'object' ? src.crop : {})
    },
    page: {
      ...base.page,
      ...(src.page && typeof src.page === 'object' ? src.page : {})
    },
    title: {
      ...base.title,
      ...(src.title && typeof src.title === 'object' ? src.title : {})
    },
    subtitle: {
      ...base.subtitle,
      ...(src.subtitle && typeof src.subtitle === 'object' ? src.subtitle : {})
    },
    grid: {
      ...base.grid,
      ...(src.grid && typeof src.grid === 'object' ? src.grid : {})
    },
    planBox: {
      ...base.planBox,
      ...(src.planBox && typeof src.planBox === 'object' ? src.planBox : {})
    },
    priceBox: {
      ...base.priceBox,
      ...(src.priceBox && typeof src.priceBox === 'object' ? src.priceBox : {})
    }
  };
  out.version = 1;
  out.crop.verticalExpandRatio = clampNumber(out.crop.verticalExpandRatio, base.crop.verticalExpandRatio, 1, 2.4);
  out.page.padding = clampNumber(out.page.padding, base.page.padding, 0, 80);
  out.title.fontSize = clampNumber(out.title.fontSize, base.title.fontSize, 12, 80);
  out.subtitle.fontSize = clampNumber(out.subtitle.fontSize, base.subtitle.fontSize, 10, 50);
  out.grid.gap = clampNumber(out.grid.gap, base.grid.gap, 0, 60);
  out.grid.leftRatio = clampNumber(out.grid.leftRatio, base.grid.leftRatio, 0.2, 4);
  out.grid.rightRatio = clampNumber(out.grid.rightRatio, base.grid.rightRatio, 0.2, 4);
  out.planBox.borderRadius = clampNumber(out.planBox.borderRadius, base.planBox.borderRadius, 0, 48);
  out.planBox.minHeight = clampNumber(out.planBox.minHeight, base.planBox.minHeight, 120, 1000);
  out.priceBox.borderRadius = clampNumber(out.priceBox.borderRadius, base.priceBox.borderRadius, 0, 48);
  out.priceBox.headingSize = clampNumber(out.priceBox.headingSize, base.priceBox.headingSize, 10, 40);
  out.priceBox.bodySize = clampNumber(out.priceBox.bodySize, base.priceBox.bodySize, 9, 32);
  out.priceBox.labelSize = clampNumber(out.priceBox.labelSize, out.priceBox.bodySize, 9, 40);
  out.priceBox.valueSize = clampNumber(out.priceBox.valueSize, out.priceBox.bodySize, 9, 40);
  out.priceBox.metaSize = clampNumber(out.priceBox.metaSize, base.priceBox.metaSize, 8, 24);
  out.priceBox.departmentLabelSize = clampNumber(out.priceBox.departmentLabelSize, out.priceBox.metaSize, 8, 40);
  out.priceBox.departmentValueSize = clampNumber(out.priceBox.departmentValueSize, out.priceBox.metaSize, 8, 40);
  out.priceBox.metersLabelSize = clampNumber(out.priceBox.metersLabelSize, out.priceBox.metaSize, 8, 40);
  out.priceBox.metersValueSize = clampNumber(out.priceBox.metersValueSize, out.priceBox.metaSize, 8, 40);
  out.priceBox.rowPaddingY = clampNumber(out.priceBox.rowPaddingY, base.priceBox.rowPaddingY, 0, 30);
  out.priceBox.lineHeight = clampNumber(out.priceBox.lineHeight, base.priceBox.lineHeight, 1, 2.4);
  out.priceBox.wordSpacing = clampNumber(out.priceBox.wordSpacing, base.priceBox.wordSpacing, 0, 20);
  out.priceBox.letterSpacing = clampNumber(out.priceBox.letterSpacing, base.priceBox.letterSpacing, 0, 10);
  out.priceBox.metaTopOffset = clampNumber(out.priceBox.metaTopOffset, base.priceBox.metaTopOffset, -30, 80);
  out.priceBox.metaGroupGap = clampNumber(out.priceBox.metaGroupGap, base.priceBox.metaGroupGap, 0, 80);
  const alignRaw = String(out.priceBox.verticalAlign || '').toLowerCase();
  out.priceBox.verticalAlign = ['start', 'center', 'end'].includes(alignRaw) ? alignRaw : base.priceBox.verticalAlign;
  if (typeof out.title.text !== 'string') out.title.text = base.title.text;
  out.title.show = Boolean(out.title.show);
  out.subtitle.show = Boolean(out.subtitle.show);
  return out;
}

function readSolarMidtownLayout() {
  const base = defaultSolarMidtownLayout();
  const repoRaw = readJson(SOLAR_MIDTOWN_LAYOUT_REPO_PATH, null);
  const seedRaw = readJson(SOLAR_MIDTOWN_LAYOUT_SEED_PATH, null);
  const runtimeRaw = readJson(SOLAR_MIDTOWN_LAYOUT_PATH, null);
  const repoNorm = normalizeSolarMidtownLayout(repoRaw);
  if (!isJsonEqual(repoNorm, normalizeSolarMidtownLayout(base))) {
    return repoNorm;
  }
  const seedNorm = normalizeSolarMidtownLayout(seedRaw);
  if (!isJsonEqual(seedNorm, normalizeSolarMidtownLayout(base))) {
    return seedNorm;
  }
  const runtimeNorm = normalizeSolarMidtownLayout(runtimeRaw);
  const merged = applyNonDefaultOverlay(
    deepCloneJson(seedNorm || repoNorm || base),
    runtimeNorm || {},
    base
  );
  return normalizeSolarMidtownLayout(merged);
}

function saveSolarMidtownLayout(layout) {
  const normalized = normalizeSolarMidtownLayout(layout);
  writeJson(SOLAR_MIDTOWN_LAYOUT_PATH, normalized);
  return normalized;
}

function defaultTablaPagosLayout() {
  return {
    version: 4,
    printScale: 0.92,
    pageMarginMm: 0,
    brandTopHeight: 226,
    railOffsetY: 226,
    railHeight: 278,
    railWidth: 54,
    logoDataUrl: '',
    logoWidth: 480,
    logoHeight: 120,
    logoOffsetX: 0,
    logoOffsetY: 0,
    railText: 'simca.mx',
    railTextSize: 18,
    titleSize: 28,
    subtitleSize: 14,
    headerPaddingTop: 12,
    headerPaddingSides: 14,
    headerPaddingBottom: 10,
    contentPaddingTop: 6,
    contentPaddingSides: 12,
    contentPaddingBottom: 12,
    contentGap: 14,
    titleOffsetX: 0,
    titleOffsetY: 0,
    showHeaderInPrint: true,
    showSummaryInPrint: true,
    showPaymentsInPrint: true,
    surfaceColor: '#e8e4da',
    paperColor: '#fffdf8',
    paperSoftColor: '#f8f2e6',
    textColor: '#1d1c18',
    titleColor: '#1d1c18',
    mutedColor: '#6f6859',
    accentColor: '#f0c419',
    railAccentColor: '#ffdb16',
    railTextColor: '#1d1c18',
    lineColor: '#ddd1bb',
    lineStrongColor: '#c7b89d',
    highlightRowColor: '#f9efbe',
    bodyFontFamily: 'Arial, sans-serif',
    headingFontFamily: 'Arial, sans-serif',
    tableFontFamily: 'Arial, sans-serif',
    summaryHeaderSize: 9,
    summaryBodySize: 12,
    paymentHeaderSize: 9,
    paymentBodySize: 12,
    paymentTitleSize: 20,
    summaryColDevelopment: 20,
    summaryColUnit: 12,
    summaryColPlan: 16,
    summaryColList: 17,
    summaryColDiscount: 17,
    summaryColPromo: 18,
    paymentsColStage: 10,
    paymentsColConcept: 22,
    paymentsColReference: 28,
    paymentsColPercent: 16,
    paymentsColAmount: 24,
    paymentsAlignStage: 'right',
    paymentsAlignConcept: 'left',
    paymentsAlignReference: 'left',
    paymentsAlignPercent: 'right',
    paymentsAlignAmount: 'right',
    lines: []
  };
}

function normalizeTablaPagosLayout(input) {
  const base = defaultTablaPagosLayout();
  const src = input && typeof input === 'object' ? input : {};
  const out = {
    ...base,
    ...src
  };
  const srcVersion = Number(src.version) || 0;
  out.version = 4;
  if (srcVersion < 3 && (!Number.isFinite(Number(src.printScale)) || Math.abs(Number(src.printScale) - 0.74) < 0.001)) {
    out.printScale = base.printScale;
  }
  if (srcVersion < 4 && (!Number.isFinite(Number(src.pageMarginMm)) || Math.abs(Number(src.pageMarginMm) - 6) < 0.001)) {
    out.pageMarginMm = base.pageMarginMm;
  }
  out.printScale = clampNumber(out.printScale, base.printScale, 0.65, 1.15);
  out.pageMarginMm = clampNumber(out.pageMarginMm, base.pageMarginMm, 0, 20);
  out.brandTopHeight = clampNumber(out.brandTopHeight, base.brandTopHeight, 120, 320);
  out.railOffsetY = clampNumber(out.railOffsetY, base.railOffsetY, 0, 500);
  out.railHeight = clampNumber(out.railHeight, base.railHeight, 40, 500);
  out.railWidth = clampNumber(out.railWidth, base.railWidth, 36, 180);
  out.logoWidth = clampNumber(out.logoWidth, base.logoWidth, 80, 900);
  out.logoHeight = clampNumber(out.logoHeight, base.logoHeight, 40, 260);
  out.logoOffsetX = clampNumber(out.logoOffsetX, base.logoOffsetX, -240, 240);
  out.logoOffsetY = clampNumber(out.logoOffsetY, base.logoOffsetY, -140, 140);
  out.railTextSize = clampNumber(out.railTextSize, base.railTextSize, 10, 28);
  out.titleSize = clampNumber(out.titleSize, base.titleSize, 16, 40);
  out.subtitleSize = clampNumber(out.subtitleSize, base.subtitleSize, 10, 22);
  out.headerPaddingTop = clampNumber(out.headerPaddingTop, base.headerPaddingTop, -160, 160);
  out.headerPaddingSides = clampNumber(out.headerPaddingSides, base.headerPaddingSides, 0, 40);
  out.headerPaddingBottom = clampNumber(out.headerPaddingBottom, base.headerPaddingBottom, 0, 40);
  out.contentPaddingTop = clampNumber(out.contentPaddingTop, base.contentPaddingTop, 0, 40);
  out.contentPaddingSides = clampNumber(out.contentPaddingSides, base.contentPaddingSides, 0, 40);
  out.contentPaddingBottom = clampNumber(out.contentPaddingBottom, base.contentPaddingBottom, 0, 40);
  out.contentGap = clampNumber(out.contentGap, base.contentGap, 0, 32);
  out.titleOffsetX = clampNumber(out.titleOffsetX, base.titleOffsetX, -240, 240);
  out.titleOffsetY = clampNumber(out.titleOffsetY, base.titleOffsetY, -120, 120);
  out.summaryHeaderSize = clampNumber(out.summaryHeaderSize, base.summaryHeaderSize, 7, 16);
  out.summaryBodySize = clampNumber(out.summaryBodySize, base.summaryBodySize, 9, 20);
  out.paymentHeaderSize = clampNumber(out.paymentHeaderSize, base.paymentHeaderSize, 7, 16);
  out.paymentBodySize = clampNumber(out.paymentBodySize, base.paymentBodySize, 9, 20);
  out.paymentTitleSize = clampNumber(out.paymentTitleSize, base.paymentTitleSize, 12, 30);
  out.showHeaderInPrint = Boolean(out.showHeaderInPrint);
  out.showSummaryInPrint = Boolean(out.showSummaryInPrint);
  out.showPaymentsInPrint = Boolean(out.showPaymentsInPrint);
  out.summaryColDevelopment = clampNumber(out.summaryColDevelopment, base.summaryColDevelopment, 10, 35);
  out.summaryColUnit = clampNumber(out.summaryColUnit, base.summaryColUnit, 8, 25);
  out.summaryColPlan = clampNumber(out.summaryColPlan, base.summaryColPlan, 10, 30);
  out.summaryColList = clampNumber(out.summaryColList, base.summaryColList, 10, 30);
  out.summaryColDiscount = clampNumber(out.summaryColDiscount, base.summaryColDiscount, 10, 30);
  out.summaryColPromo = clampNumber(out.summaryColPromo, base.summaryColPromo, 10, 30);
  out.paymentsColStage = clampNumber(out.paymentsColStage, base.paymentsColStage, 6, 24);
  out.paymentsColConcept = clampNumber(out.paymentsColConcept, base.paymentsColConcept, 10, 36);
  out.paymentsColReference = clampNumber(out.paymentsColReference, base.paymentsColReference, 12, 40);
  out.paymentsColPercent = clampNumber(out.paymentsColPercent, base.paymentsColPercent, 8, 30);
  out.paymentsColAmount = clampNumber(out.paymentsColAmount, base.paymentsColAmount, 10, 30);
  const validAlign = ['left', 'center', 'right'];
  out.paymentsAlignStage = validAlign.includes(String(out.paymentsAlignStage || '').toLowerCase()) ? String(out.paymentsAlignStage).toLowerCase() : base.paymentsAlignStage;
  out.paymentsAlignConcept = validAlign.includes(String(out.paymentsAlignConcept || '').toLowerCase()) ? String(out.paymentsAlignConcept).toLowerCase() : base.paymentsAlignConcept;
  out.paymentsAlignReference = validAlign.includes(String(out.paymentsAlignReference || '').toLowerCase()) ? String(out.paymentsAlignReference).toLowerCase() : base.paymentsAlignReference;
  out.paymentsAlignPercent = validAlign.includes(String(out.paymentsAlignPercent || '').toLowerCase()) ? String(out.paymentsAlignPercent).toLowerCase() : base.paymentsAlignPercent;
  out.paymentsAlignAmount = validAlign.includes(String(out.paymentsAlignAmount || '').toLowerCase()) ? String(out.paymentsAlignAmount).toLowerCase() : base.paymentsAlignAmount;
  if (typeof out.logoDataUrl !== 'string') out.logoDataUrl = base.logoDataUrl;
  if (typeof out.railText !== 'string') out.railText = base.railText;
  if (typeof out.surfaceColor !== 'string') out.surfaceColor = base.surfaceColor;
  if (typeof out.paperColor !== 'string') out.paperColor = base.paperColor;
  if (typeof out.paperSoftColor !== 'string') out.paperSoftColor = base.paperSoftColor;
  if (typeof out.textColor !== 'string') out.textColor = base.textColor;
  if (typeof out.titleColor !== 'string') out.titleColor = base.titleColor;
  if (typeof out.mutedColor !== 'string') out.mutedColor = base.mutedColor;
  if (typeof out.accentColor !== 'string') out.accentColor = base.accentColor;
  if (typeof out.railAccentColor !== 'string') out.railAccentColor = base.railAccentColor;
  if (typeof out.railTextColor !== 'string') out.railTextColor = base.railTextColor;
  if (typeof out.lineColor !== 'string') out.lineColor = base.lineColor;
  if (typeof out.lineStrongColor !== 'string') out.lineStrongColor = base.lineStrongColor;
  if (typeof out.highlightRowColor !== 'string') out.highlightRowColor = base.highlightRowColor;
  if (typeof out.bodyFontFamily !== 'string') out.bodyFontFamily = base.bodyFontFamily;
  if (typeof out.headingFontFamily !== 'string') out.headingFontFamily = base.headingFontFamily;
  if (typeof out.tableFontFamily !== 'string') out.tableFontFamily = base.tableFontFamily;
  out.lines = Array.isArray(out.lines)
    ? out.lines.slice(0, 24).map((line, index) => {
        const srcLine = line && typeof line === 'object' ? line : {};
        return {
          id: typeof srcLine.id === 'string' && srcLine.id ? srcLine.id : `line-${index + 1}`,
          orientation: String(srcLine.orientation || 'horizontal').toLowerCase() === 'vertical' ? 'vertical' : 'horizontal',
          x: clampNumber(srcLine.x, srcLine.x1, 0, 100),
          y: clampNumber(srcLine.y, srcLine.y1, 0, 100),
          length: clampNumber(srcLine.length, String(srcLine.orientation || 'horizontal').toLowerCase() === 'vertical'
            ? Math.max(0, (Number(srcLine.y2) || 0) - (Number(srcLine.y1) || 0))
            : Math.max(0, (Number(srcLine.x2) || 100) - (Number(srcLine.x1) || 0)), 0, 100),
          width: clampNumber(srcLine.width, 1, 1, 12),
          color: typeof srcLine.color === 'string' && srcLine.color ? srcLine.color : base.lineStrongColor,
          opacity: clampNumber(srcLine.opacity, 1, 0.1, 1)
        };
      })
    : [];
  return out;
}

function readTablaPagosLayout() {
  return normalizeTablaPagosLayout(readJson(TABLA_PAGOS_LAYOUT_PATH, null));
}

function saveTablaPagosLayout(layout) {
  const normalized = normalizeTablaPagosLayout(layout);
  writeJson(TABLA_PAGOS_LAYOUT_PATH, normalized);
  return normalized;
}

function defaultViceroyPaymentPlanLayout() {
  return {
    version: 1,
    printScale: 0.92,
    pageMarginMm: 6,
    heroPaddingTop: 18,
    heroPaddingSides: 18,
    heroPaddingBottom: 14,
    contentPaddingTop: 10,
    contentPaddingSides: 14,
    contentPaddingBottom: 14,
    contentGap: 10,
    headlineSize: 15,
    summaryHeaderSize: 9,
    summaryValueSize: 16,
    paymentsTitleSize: 15,
    paymentsHeaderSize: 9,
    paymentsBodySize: 11,
    summaryRowPaddingY: 8,
    paymentRowPaddingY: 6,
    oliveLineColor: '#9b9b8b',
    oliveLineWidth: 1,
    shellBorderColor: '#d6d0c2',
    shellBorderWidth: 1,
    bodyFontFamily: 'Arial, sans-serif',
    headingFontFamily: 'Georgia, "Times New Roman", serif',
    tableFontFamily: 'Arial, sans-serif',
    showHeroInPrint: true,
    showLogosInPrint: false,
    showSummaryInPrint: true,
    showPaymentsInPrint: true,
    paymentsColStage: 10,
    paymentsColConcept: 22,
    paymentsColReference: 22,
    paymentsColPercent: 18,
    paymentsColAmount: 18,
    paymentsAlignStage: 'right',
    paymentsAlignConcept: 'left',
    paymentsAlignReference: 'left',
    paymentsAlignPercent: 'right',
    paymentsAlignAmount: 'right',
    summaryColUnit: 14,
    summaryColPlan: 26,
    summaryColList: 20,
    summaryColDiscount: 20,
    summaryColPromo: 20
  };
}

function normalizeViceroyPaymentPlanLayout(input) {
  const base = defaultViceroyPaymentPlanLayout();
  const src = input && typeof input === 'object' ? input : {};
  const out = {
    ...base,
    ...src
  };
  out.version = 1;
  out.printScale = clampNumber(out.printScale, base.printScale, 0.65, 1.15);
  out.pageMarginMm = clampNumber(out.pageMarginMm, base.pageMarginMm, 0, 20);
  out.heroPaddingTop = clampNumber(out.heroPaddingTop, base.heroPaddingTop, 0, 60);
  out.heroPaddingSides = clampNumber(out.heroPaddingSides, base.heroPaddingSides, 0, 60);
  out.heroPaddingBottom = clampNumber(out.heroPaddingBottom, base.heroPaddingBottom, 0, 60);
  out.contentPaddingTop = clampNumber(out.contentPaddingTop, base.contentPaddingTop, 0, 40);
  out.contentPaddingSides = clampNumber(out.contentPaddingSides, base.contentPaddingSides, 0, 40);
  out.contentPaddingBottom = clampNumber(out.contentPaddingBottom, base.contentPaddingBottom, 0, 40);
  out.contentGap = clampNumber(out.contentGap, base.contentGap, 0, 32);
  out.headlineSize = clampNumber(out.headlineSize, base.headlineSize, 10, 28);
  out.summaryHeaderSize = clampNumber(out.summaryHeaderSize, base.summaryHeaderSize, 7, 16);
  out.summaryValueSize = clampNumber(out.summaryValueSize, base.summaryValueSize, 10, 30);
  out.paymentsTitleSize = clampNumber(out.paymentsTitleSize, base.paymentsTitleSize, 10, 26);
  out.paymentsHeaderSize = clampNumber(out.paymentsHeaderSize, base.paymentsHeaderSize, 7, 16);
  out.paymentsBodySize = clampNumber(out.paymentsBodySize, base.paymentsBodySize, 8, 18);
  out.summaryRowPaddingY = clampNumber(out.summaryRowPaddingY, base.summaryRowPaddingY, 2, 20);
  out.paymentRowPaddingY = clampNumber(out.paymentRowPaddingY, base.paymentRowPaddingY, 2, 20);
  out.oliveLineWidth = clampNumber(out.oliveLineWidth, base.oliveLineWidth, 0, 4);
  out.shellBorderWidth = clampNumber(out.shellBorderWidth, base.shellBorderWidth, 0, 4);
  if (typeof out.bodyFontFamily !== 'string') out.bodyFontFamily = base.bodyFontFamily;
  if (typeof out.headingFontFamily !== 'string') out.headingFontFamily = base.headingFontFamily;
  if (typeof out.tableFontFamily !== 'string') out.tableFontFamily = base.tableFontFamily;
  out.showHeroInPrint = Boolean(out.showHeroInPrint);
  out.showLogosInPrint = Boolean(out.showLogosInPrint);
  out.showSummaryInPrint = Boolean(out.showSummaryInPrint);
  out.showPaymentsInPrint = Boolean(out.showPaymentsInPrint);
  out.paymentsColStage = clampNumber(out.paymentsColStage, base.paymentsColStage, 6, 24);
  out.paymentsColConcept = clampNumber(out.paymentsColConcept, base.paymentsColConcept, 10, 36);
  out.paymentsColReference = clampNumber(out.paymentsColReference, base.paymentsColReference, 10, 36);
  out.paymentsColPercent = clampNumber(out.paymentsColPercent, base.paymentsColPercent, 8, 30);
  out.paymentsColAmount = clampNumber(out.paymentsColAmount, base.paymentsColAmount, 8, 30);
  out.summaryColUnit = clampNumber(out.summaryColUnit, base.summaryColUnit, 8, 40);
  out.summaryColPlan = clampNumber(out.summaryColPlan, base.summaryColPlan, 12, 45);
  out.summaryColList = clampNumber(out.summaryColList, base.summaryColList, 10, 35);
  out.summaryColDiscount = clampNumber(out.summaryColDiscount, base.summaryColDiscount, 10, 35);
  out.summaryColPromo = clampNumber(out.summaryColPromo, base.summaryColPromo, 10, 35);
  const validAlign = ['left', 'center', 'right'];
  out.paymentsAlignStage = validAlign.includes(String(out.paymentsAlignStage || '').toLowerCase()) ? String(out.paymentsAlignStage).toLowerCase() : base.paymentsAlignStage;
  out.paymentsAlignConcept = validAlign.includes(String(out.paymentsAlignConcept || '').toLowerCase()) ? String(out.paymentsAlignConcept).toLowerCase() : base.paymentsAlignConcept;
  out.paymentsAlignReference = validAlign.includes(String(out.paymentsAlignReference || '').toLowerCase()) ? String(out.paymentsAlignReference).toLowerCase() : base.paymentsAlignReference;
  out.paymentsAlignPercent = validAlign.includes(String(out.paymentsAlignPercent || '').toLowerCase()) ? String(out.paymentsAlignPercent).toLowerCase() : base.paymentsAlignPercent;
  out.paymentsAlignAmount = validAlign.includes(String(out.paymentsAlignAmount || '').toLowerCase()) ? String(out.paymentsAlignAmount).toLowerCase() : base.paymentsAlignAmount;
  if (typeof out.oliveLineColor !== 'string') out.oliveLineColor = base.oliveLineColor;
  if (typeof out.shellBorderColor !== 'string') out.shellBorderColor = base.shellBorderColor;
  return out;
}

function readViceroyPaymentPlanLayout() {
  return normalizeViceroyPaymentPlanLayout(readJson(VICEROY_PAYMENT_PLAN_LAYOUT_PATH, null));
}

function saveViceroyPaymentPlanLayout(layout) {
  const normalized = normalizeViceroyPaymentPlanLayout(layout);
  writeJson(VICEROY_PAYMENT_PLAN_LAYOUT_PATH, normalized);
  return normalized;
}

function defaultBrokersSimcaTemplate() {
  return {
    version: 2,
    sourceHtml: `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Documento editable</title>
  <style>
    body{margin:0;font-family:Arial,sans-serif;background:#f4efe6;color:#1b1b1b}
    .sheet{max-width:980px;margin:24px auto;padding:32px;background:#fff;border-radius:18px;box-shadow:0 18px 50px rgba(0,0,0,.08)}
    .hero{padding:56px 40px;border-radius:16px;background:linear-gradient(135deg,#e6ded0,#f7f3eb)}
    h1{margin:0 0 12px;font-size:42px}
    p{line-height:1.6}
    .grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px;margin-top:24px}
    .card{padding:20px;border:1px solid #e7dfd0;border-radius:14px;background:#fcfaf6}
    @media (max-width:720px){.sheet{margin:0;border-radius:0;padding:18px}.hero{padding:28px 20px}.grid{grid-template-columns:1fr}h1{font-size:30px}}
  </style>
</head>
<body>
  <main class="sheet">
    <section class="hero">
      <h1>Documento editable</h1>
      <p>Pega aqui tu propio HTML, luego selecciona elementos en el preview para cambiar texto, estilos y posicion.</p>
    </section>
    <section class="grid">
      <article class="card">
        <h2>Bloque 1</h2>
        <p>Este es un contenido de muestra para que el editor no abra vacio.</p>
      </article>
      <article class="card">
        <h2>Bloque 2</h2>
        <p>Cuando pegues tu HTML, esta plantilla base se reemplaza completa.</p>
      </article>
    </section>
  </main>
</body>
</html>`
  };
}

function normalizeBrokersSimcaTemplate(input) {
  const base = defaultBrokersSimcaTemplate();
  const src = isPlainObject(input) ? input : {};
  const sourceHtml = typeof src.sourceHtml === 'string' && src.sourceHtml.trim()
    ? src.sourceHtml
    : base.sourceHtml;
  return {
    version: 2,
    sourceHtml
  };
}

function readBrokersSimcaTemplate() {
  return normalizeBrokersSimcaTemplate(readJson(BROKERS_SIMCA_TEMPLATE_PATH, null));
}

function saveBrokersSimcaTemplate(template) {
  const normalized = normalizeBrokersSimcaTemplate(template);
  writeJson(BROKERS_SIMCA_TEMPLATE_PATH, normalized);
  return normalized;
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

function normalizeWhisperlistCountry(raw) {
  const value = String(raw || '').trim().toUpperCase().replace(/\s+/g, ' ');
  if (!value) return '';
  if (value === 'MEXICO' || value === 'MÉXICO') return 'MEXICO';
  if (value === 'USA' || value === 'US' || value === 'ESTADOS UNIDOS' || value === 'EEUU' || value === 'EUA') return 'USA';
  if (value === 'CANADA' || value === 'CANADÁ') return 'CANADA';
  if (value === 'EUROPA') return 'EUROPA';
  if (value === 'RESTO AMERICA' || value === 'RESTO DE AMERICA' || value === 'RESTO DE AMÉRICA') return 'RESTO AMERICA';
  if (value === 'RESTO CONTINENTES' || value === 'RESTO DE CONTINENTES' || value === 'RESTO DE LOS CONTINENTES') return '';
  return WHISPERLIST_PAISES.includes(value) ? value : '';
}

function normalizeWhisperlistCity(raw) {
  return String(raw || '')
    .trim()
    .replace(/\s+/g, ' ');
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

function normalizeYesNo(raw) {
  const value = String(raw || '').trim().toUpperCase();
  if (['SI', 'YES', 'Y', 'TRUE', '1'].includes(value)) return 'SI';
  if (['NO', 'N', 'FALSE', '0'].includes(value)) return 'NO';
  return '';
}

function normalizeWhisperlistKpi(raw) {
  const item = raw && typeof raw === 'object' ? raw : {};
  return {
    presupuesto: String(item.presupuesto || '').trim(),
    vista: ['32', '38', 'MAR', 'INTERIOR'].includes(String(item.vista || '').trim().toUpperCase())
      ? String(item.vista || '').trim().toUpperCase()
      : '',
    prioridades: String(item.prioridades || '').trim(),
    opcionA: String(item.opcionA || '').trim(),
    opcionB: String(item.opcionB || '').trim(),
    opcionC: String(item.opcionC || '').trim(),
    opcionD: String(item.opcionD || '').trim(),
    hojaReserva: normalizeYesNo(item.hojaReserva) || 'NO',
    reservaPagada: normalizeYesNo(item.reservaPagada) || 'NO',
    unidadAsignada: normalizeYesNo(item.unidadAsignada) || 'NO',
    contratoEnviado: normalizeYesNo(item.contratoEnviado) || 'NO',
    contratoFirmado: normalizeYesNo(item.contratoFirmado) || 'NO',
    enganchePagado: normalizeYesNo(item.enganchePagado) || 'NO'
  };
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
  const correo = String(
    normalized.correo
    || normalized.correo_asesor
    || normalized.email_asesor
    || ''
  ).trim().toLowerCase();
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
    pais: normalizeWhisperlistCountry(normalized.pais || normalized.country || normalized.pais_cliente),
    ciudad: normalizeWhisperlistCity(normalized.ciudad || normalized.city || normalized.ciudad_cliente),
    clientEmail: normalizeClientEmail(normalized.correo_cliente || normalized.email_cliente || normalized.client_email),
    clientPhone: normalizeClientPhone(normalized.telefono_cliente || normalized.telefono || normalized.client_phone),
    kpi: normalizeWhisperlistKpi({
      presupuesto: normalized.presupuesto || normalized.kpi_presupuesto,
      vista: normalized.vista || normalized.kpi_vista,
      prioridades: normalized.prioridades || normalized.kpi_prioridades,
      opcionA: normalized.opcion_a || normalized.kpi_opcion_a,
      opcionB: normalized.opcion_b || normalized.kpi_opcion_b,
      opcionC: normalized.opcion_c || normalized.kpi_opcion_c,
      opcionD: normalized.opcion_d || normalized.kpi_opcion_d,
      hojaReserva: normalized.hoja_de_reserva || normalized.kpi_hoja_de_reserva,
      reservaPagada: normalized.reserva_pagada || normalized.kpi_reserva_pagada,
      unidadAsignada: normalized.unidad_asignada || normalized.kpi_unidad_asignada,
      contratoEnviado: normalized.contrato_enviado || normalized.kpi_contrato_enviado,
      contratoFirmado: normalized.contrato_firmado || normalized.kpi_contrato_firmado,
      enganchePagado: normalized.enganche_pagado || normalized.kpi_enganche_pagado
    }),
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
      pais TEXT NOT NULL DEFAULT '',
      ciudad TEXT NOT NULL DEFAULT '',
      client_email TEXT NOT NULL DEFAULT '',
      client_phone TEXT NOT NULL DEFAULT '',
      kpi_json TEXT NOT NULL DEFAULT '',
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )
  `);
  await whisperlistPool.query(`ALTER TABLE whisperlist_rows ADD COLUMN IF NOT EXISTS recamaras TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE whisperlist_rows ADD COLUMN IF NOT EXISTS pais TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE whisperlist_rows ADD COLUMN IF NOT EXISTS ciudad TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE whisperlist_rows ADD COLUMN IF NOT EXISTS client_email TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE whisperlist_rows ADD COLUMN IF NOT EXISTS client_phone TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE whisperlist_rows ADD COLUMN IF NOT EXISTS kpi_json TEXT NOT NULL DEFAULT ''`);
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
      SELECT id, asesor, correo, canal, tipo_venta, nombre_cliente, recamaras, pais, ciudad, client_email, client_phone, kpi_json, updated_at
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
      pais: String(row.pais || ''),
      ciudad: String(row.ciudad || ''),
      clientEmail: String(row.client_email || ''),
      clientPhone: String(row.client_phone || ''),
      kpi: (() => {
        try {
          return normalizeWhisperlistKpi(row.kpi_json ? JSON.parse(String(row.kpi_json)) : {});
        } catch {
          return normalizeWhisperlistKpi({});
        }
      })(),
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
    pais: normalizeWhisperlistCountry(row.pais),
    ciudad: normalizeWhisperlistCity(row.ciudad),
    canal: normalizeWhisperlistCanal(row.canal),
    tipoVenta: normalizeWhisperlistTipoVenta(row.tipoVenta),
    kpi: normalizeWhisperlistKpi(row.kpi)
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
        `INSERT INTO whisperlist_rows (id, asesor, correo, canal, tipo_venta, nombre_cliente, recamaras, pais, ciudad, client_email, client_phone, kpi_json, updated_at)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13)`,
        [
          String(row.id || `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`),
          normalizeWhisperlistAsesor(row.asesor),
          String(row.correo || '').trim().toLowerCase(),
          String(row.canal || '').trim(),
          String(row.tipoVenta || '').trim(),
          normalizeWhisperlistPersonText(row.nombreCliente),
          normalizeWhisperlistRecamaras(row.recamaras),
          normalizeWhisperlistCountry(row.pais),
          normalizeWhisperlistCity(row.ciudad),
          normalizeClientEmail(row.clientEmail),
          normalizeClientPhone(row.clientPhone),
          JSON.stringify(normalizeWhisperlistKpi(row.kpi)),
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

const VICEROY_ROOM_OPTIONS = new Set(['sala-grande', 'sala-chica']);
const VICEROY_ROOM_HOURS = new Set(['10:00', '11:00', '12:00', '13:00', '14:00', '15:00', '16:00', '17:00', '18:00']);

function normalizeReservationRoom(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (VICEROY_ROOM_OPTIONS.has(normalized)) return normalized;
  return '';
}

function normalizeReservationDate(value) {
  const normalized = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(normalized) ? normalized : '';
}

function normalizeReservationHour(value) {
  const normalized = String(value || '').trim();
  return VICEROY_ROOM_HOURS.has(normalized) ? normalized : '';
}

function normalizeReservationTitle(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().slice(0, 160);
}

function normalizeRoomReservationRow(row) {
  return {
    id: String(row && row.id || '').trim(),
    date: normalizeReservationDate(row && row.date),
    room: normalizeReservationRoom(row && row.room),
    hour: normalizeReservationHour(row && row.hour),
    title: normalizeReservationTitle(row && row.title),
    createdByEmail: String(row && row.createdByEmail || '').trim().toLowerCase(),
    createdByName: String(row && row.createdByName || '').trim(),
    updatedAt: String(row && row.updatedAt || '').trim() || new Date().toISOString()
  };
}

function sortRoomReservationRows(rows) {
  return [...(Array.isArray(rows) ? rows : [])].sort((a, b) => {
    const byDate = String(a.date || '').localeCompare(String(b.date || ''));
    if (byDate !== 0) return byDate;
    const byRoom = String(a.room || '').localeCompare(String(b.room || ''));
    if (byRoom !== 0) return byRoom;
    const byHour = String(a.hour || '').localeCompare(String(b.hour || ''));
    if (byHour !== 0) return byHour;
    return String(a.title || '').localeCompare(String(b.title || ''), 'es');
  });
}

function readViceroyRoomReservations() {
  const raw = readJson(VICEROY_ROOM_RESERVATIONS_PATH, { rows: [], updatedAt: null });
  const rows = sortRoomReservationRows(
    (Array.isArray(raw.rows) ? raw.rows : [])
      .map((row) => normalizeRoomReservationRow(row))
      .filter((row) => row.date && row.room && row.hour && row.title)
  );
  return {
    rows,
    updatedAt: raw.updatedAt || null
  };
}

function saveViceroyRoomReservations(rows) {
  const normalizedRows = sortRoomReservationRows(
    (Array.isArray(rows) ? rows : [])
      .map((row) => normalizeRoomReservationRow(row))
      .filter((row) => row.date && row.room && row.hour && row.title)
  );
  writeJson(VICEROY_ROOM_RESERVATIONS_PATH, {
    rows: normalizedRows,
    updatedAt: new Date().toISOString()
  });
}

async function isAllowedLoginEmail(email) {
  const normalized = String(email || '').trim().toLowerCase();
  if (!normalized) return false;
  if (normalized === GERENTE_EMAIL) return true;
  if (normalized.endsWith(`@${ALLOWED_DOMAIN}`)) return true;
  if (EXTRA_ALLOWED_EMAILS.has(normalized)) return true;
  if ((await whisperlistAllowedEmails()).has(normalized)) return true;
  if ((await viceroyRegistrosAllowedEmails()).has(normalized)) return true;
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

function splitIntoChunks(listInput, chunkSizeInput) {
  const list = Array.isArray(listInput) ? listInput : [];
  const chunkSize = Math.max(1, Math.floor(Number(chunkSizeInput) || 1));
  const out = [];
  for (let i = 0; i < list.length; i += chunkSize) {
    out.push(list.slice(i, i + chunkSize));
  }
  return out;
}

async function mapWithConcurrency(itemsInput, concurrencyInput, worker) {
  const items = Array.isArray(itemsInput) ? itemsInput : [];
  const concurrency = Math.max(1, Math.floor(Number(concurrencyInput) || 1));
  const results = new Array(items.length);
  let idx = 0;
  async function run() {
    while (true) {
      const current = idx;
      idx += 1;
      if (current >= items.length) return;
      results[current] = await worker(items[current], current);
    }
  }
  const runners = Array.from({ length: Math.min(concurrency, items.length) }, () => run());
  await Promise.all(runners);
  return results;
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
    PAIS: normalizeWhisperlistCountry(row.pais),
    CIUDAD: normalizeWhisperlistCity(row.ciudad),
    CORREO_CLIENTE: normalizeClientEmail(row.clientEmail),
    TELEFONO_CLIENTE: normalizeClientPhone(row.clientPhone),
    KPI_PRESUPUESTO: String(row && row.kpi && row.kpi.presupuesto || ''),
    KPI_VISTA: String(row && row.kpi && row.kpi.vista || ''),
    KPI_PRIORIDADES: String(row && row.kpi && row.kpi.prioridades || ''),
    KPI_OPCION_A: String(row && row.kpi && row.kpi.opcionA || ''),
    KPI_OPCION_B: String(row && row.kpi && row.kpi.opcionB || ''),
    KPI_OPCION_C: String(row && row.kpi && row.kpi.opcionC || ''),
    KPI_OPCION_D: String(row && row.kpi && row.kpi.opcionD || ''),
    KPI_HOJA_DE_RESERVA: normalizeYesNo(row && row.kpi && row.kpi.hojaReserva),
    KPI_RESERVA_PAGADA: normalizeYesNo(row && row.kpi && row.kpi.reservaPagada),
    KPI_UNIDAD_ASIGNADA: normalizeYesNo(row && row.kpi && row.kpi.unidadAsignada),
    KPI_CONTRATO_ENVIADO: normalizeYesNo(row && row.kpi && row.kpi.contratoEnviado),
    KPI_CONTRATO_FIRMADO: normalizeYesNo(row && row.kpi && row.kpi.contratoFirmado),
    KPI_ENGANCHE_PAGADO: normalizeYesNo(row && row.kpi && row.kpi.enganchePagado),
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

function normalizeViceroyRegistrosRow(rawRow, fallbackId) {
  const normalized = {};
  Object.entries(rawRow || {}).forEach(([key, value]) => {
    normalized[normalizeWhisperlistKey(key)] = value;
  });

  const asesor = normalizeWhisperlistAsesor(normalized.asesor);
  const correo = String(
    normalized.correo
    || normalized.correo_asesor
    || normalized.email_asesor
    || ''
  ).trim().toLowerCase();
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
    pais: normalizeWhisperlistCountry(normalized.pais || normalized.country || normalized.pais_cliente),
    ciudad: normalizeWhisperlistCity(normalized.ciudad || normalized.city || normalized.ciudad_cliente),
    clientEmail: normalizeClientEmail(normalized.correo_cliente || normalized.email_cliente || normalized.client_email),
    clientPhone: normalizeClientPhone(normalized.telefono_cliente || normalized.telefono || normalized.client_phone),
    updatedAt: new Date().toISOString()
  };
}

function parseViceroyRegistrosWorkbook(workbook, requestedSheetName) {
  const fallbackSheet = workbook.SheetNames.includes('Hoja1')
    ? 'Hoja1'
    : (workbook.SheetNames[0] || '');
  const sheetName = String(requestedSheetName || fallbackSheet || '').trim();
  if (!sheetName) return { sheetName: '', rows: [] };
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return { sheetName: '', rows: [] };
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  const rows = rawRows.map((item) => normalizeViceroyRegistrosRow(item)).filter(Boolean);
  return { sheetName, rows };
}

async function ensureViceroyRegistrosDbSchema() {
  if (!whisperlistPool) return;
  await whisperlistPool.query(`
    CREATE TABLE IF NOT EXISTS viceroy_registros_rows (
      id TEXT PRIMARY KEY,
      asesor TEXT NOT NULL,
      correo TEXT NOT NULL,
      canal TEXT NOT NULL DEFAULT '',
      tipo_venta TEXT NOT NULL DEFAULT '',
      nombre_cliente TEXT NOT NULL DEFAULT '',
      recamaras TEXT NOT NULL DEFAULT '',
      pais TEXT NOT NULL DEFAULT '',
      ciudad TEXT NOT NULL DEFAULT '',
      client_email TEXT NOT NULL DEFAULT '',
      client_phone TEXT NOT NULL DEFAULT '',
      updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )
  `);
  await whisperlistPool.query(`ALTER TABLE viceroy_registros_rows ADD COLUMN IF NOT EXISTS tipo_venta TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE viceroy_registros_rows ADD COLUMN IF NOT EXISTS recamaras TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE viceroy_registros_rows ADD COLUMN IF NOT EXISTS pais TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE viceroy_registros_rows ADD COLUMN IF NOT EXISTS ciudad TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE viceroy_registros_rows ADD COLUMN IF NOT EXISTS client_email TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`ALTER TABLE viceroy_registros_rows ADD COLUMN IF NOT EXISTS client_phone TEXT NOT NULL DEFAULT ''`);
  await whisperlistPool.query(`
    CREATE TABLE IF NOT EXISTS viceroy_registros_meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    )
  `);
}

async function readViceroyRegistrosData() {
  if (!viceroyRegistrosStorageReady) await ensureViceroyRegistrosStorageReady();
  if (!whisperlistPool) {
    const raw = readJson(VICEROY_REGISTROS_JSON_PATH, { rows: [], updatedAt: null, sourceFile: '' });
    return {
      rows: sortWhisperlistRows(Array.isArray(raw.rows) ? raw.rows : []),
      updatedAt: raw.updatedAt || null,
      sourceFile: String(raw.sourceFile || '')
    };
  }

  const [rowsRes, metaRes] = await Promise.all([
    whisperlistPool.query(`
      SELECT id, asesor, correo, canal, tipo_venta, nombre_cliente, recamaras, pais, ciudad, client_email, client_phone, updated_at
      FROM viceroy_registros_rows
      ORDER BY updated_at DESC, id ASC
    `),
    whisperlistPool.query(`
      SELECT key, value
      FROM viceroy_registros_meta
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
      pais: String(row.pais || ''),
      ciudad: String(row.ciudad || ''),
      clientEmail: String(row.client_email || ''),
      clientPhone: String(row.client_phone || ''),
      updatedAt: row.updated_at ? new Date(row.updated_at).toISOString() : new Date().toISOString()
    }))),
    updatedAt: meta.updatedAt || null,
    sourceFile: meta.sourceFile || ''
  };
}

async function saveViceroyRegistrosRows(rows, sourceFile) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const normalizedRows = safeRows.map((row) => ({
    ...row,
    asesor: normalizeWhisperlistAsesor(row.asesor),
    nombreCliente: normalizeWhisperlistPersonText(row.nombreCliente),
    correo: String(row.correo || '').trim().toLowerCase(),
    clientEmail: normalizeClientEmail(row.clientEmail),
    clientPhone: normalizeClientPhone(row.clientPhone),
    recamaras: normalizeWhisperlistRecamaras(row.recamaras),
    pais: normalizeWhisperlistCountry(row.pais),
    ciudad: normalizeWhisperlistCity(row.ciudad),
    canal: normalizeWhisperlistCanal(row.canal),
    tipoVenta: normalizeWhisperlistTipoVenta(row.tipoVenta)
  }));
  const updatedAt = new Date().toISOString();
  if (!whisperlistPool) {
    writeJson(VICEROY_REGISTROS_JSON_PATH, {
      rows: normalizedRows,
      updatedAt,
      sourceFile: String(sourceFile || '')
    });
    return;
  }

  const client = await whisperlistPool.connect();
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM viceroy_registros_rows');
    for (const row of normalizedRows) {
      await client.query(
        `INSERT INTO viceroy_registros_rows (id, asesor, correo, canal, tipo_venta, nombre_cliente, recamaras, pais, ciudad, client_email, client_phone, updated_at)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12)`,
        [
          String(row.id || `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`),
          normalizeWhisperlistAsesor(row.asesor),
          String(row.correo || '').trim().toLowerCase(),
          String(row.canal || '').trim(),
          String(row.tipoVenta || '').trim(),
          normalizeWhisperlistPersonText(row.nombreCliente),
          normalizeWhisperlistRecamaras(row.recamaras),
          normalizeWhisperlistCountry(row.pais),
          normalizeWhisperlistCity(row.ciudad),
          normalizeClientEmail(row.clientEmail),
          normalizeClientPhone(row.clientPhone),
          String(row.updatedAt || updatedAt)
        ]
      );
    }
    await client.query(
      `INSERT INTO viceroy_registros_meta (key, value)
       VALUES ('sourceFile', $1)
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
      [String(sourceFile || '')]
    );
    await client.query(
      `INSERT INTO viceroy_registros_meta (key, value)
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

function viceroyRegistrosExportRows(rows) {
  const items = Array.isArray(rows) ? rows : [];
  return items.map((row) => ({
    ASESOR: normalizeWhisperlistAsesor(row.asesor),
    CORREO_ASESOR: String(row.correo || '').trim().toLowerCase(),
    CANAL: normalizeWhisperlistCanal(row.canal),
    TIPO_DE_VENTA: normalizeWhisperlistTipoVenta(row.tipoVenta),
    NOMBRE_CLIENTE: normalizeWhisperlistPersonText(row.nombreCliente),
    RECAMARAS: normalizeWhisperlistRecamaras(row.recamaras),
    PAIS: normalizeWhisperlistCountry(row.pais),
    CIUDAD: normalizeWhisperlistCity(row.ciudad),
    CORREO_CLIENTE: normalizeClientEmail(row.clientEmail),
    TELEFONO_CLIENTE: normalizeClientPhone(row.clientPhone),
    UPDATED_AT: String(row.updatedAt || '')
  }));
}

async function viceroyRegistrosAllowedEmails() {
  const data = await readViceroyRegistrosData();
  const emails = new Set();
  data.rows.forEach((row) => {
    const email = String(row && row.correo || '').trim().toLowerCase();
    if (email) emails.add(email);
  });
  return emails;
}

async function seedViceroyRegistrosFromExcelIfNeeded() {
  if (!fs.existsSync(VICEROY_REGISTROS_EXCEL_PATH)) return;
  if (!whisperlistPool && fs.existsSync(VICEROY_REGISTROS_JSON_PATH)) return;
  try {
    if (whisperlistPool) {
      const countRes = await whisperlistPool.query('SELECT COUNT(*)::int AS total FROM viceroy_registros_rows');
      const count = Number(countRes.rows[0] && countRes.rows[0].total || 0);
      if (count > 0) return;
    }
    const workbook = XLSX.readFile(VICEROY_REGISTROS_EXCEL_PATH, { cellDates: true });
    const parsed = parseViceroyRegistrosWorkbook(workbook);
    await saveViceroyRegistrosRows(parsed.rows, path.basename(VICEROY_REGISTROS_EXCEL_PATH));
    log(`Viceroy Registros inicializado con ${parsed.rows.length} filas`);
  } catch (err) {
    log(`No se pudo inicializar Viceroy Registros: ${err && err.message ? err.message : err}`);
  }
}

async function ensureViceroyRegistrosStorageReady() {
  if (viceroyRegistrosStorageReady) return;
  log(`Viceroy Registros storage: ${whisperlistPool ? 'postgres' : 'json-file'}`);
  if (whisperlistPool) await ensureViceroyRegistrosDbSchema();
  await seedViceroyRegistrosFromExcelIfNeeded();
  viceroyRegistrosStorageReady = true;
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
  planLink: ['link', 'plano', 'plan_link', 'planlink', 'url_plano', 'urlplano', 'plano_url', 'image_url', 'imageurl'],
  recamaras: ['recamaras', 'recamaras_', 'rec', 'bedrooms', 'beds', 'habitaciones'],
  building: ['edificio', 'building', 'torre', 'tower', 'fase', 'phase'],
  tipologia: ['tipologia', 'tipologia_', 'typology', 'tipo', 'tipo_unidad', 'type'],
  view: ['vista', 'view', 'vistas'],
  assignment: ['asignacion', 'assignment', 'assigned_to', 'assignedto', 'asesor', 'advisor', 'broker'],
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

function normalizeViceroyPlanLink(rawValue) {
  const value = String(rawValue == null ? '' : rawValue).trim();
  if (!value) return '';
  if (/^https?:\/\//i.test(value)) return value;
  return '';
}

function normalizeViceroyInventoryRow(rawRow) {
  const row = {};
  Object.entries(rawRow || {}).forEach(([key, value]) => {
    row[normalizeHeaderKey(key)] = value;
  });
  const unit = String(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.unit) || '').trim();
  if (!unit) return null;
  const development = String(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.development) || '').trim();
  const planLink = normalizeViceroyPlanLink(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.planLink));
  const building = String(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.building) || '').trim();
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
  const assignment = String(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.assignment) || '').trim();
  const m2Raw = pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.m2);
  const sqftRaw = pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.sqft);
  const priceRaw = pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.price);
  const status = normalizeViceroyRowStatus(pickValueByAliases(row, VICEROY_PILOTO_COLUMN_ALIASES.status));
  return {
    development,
    unidad: unit,
    planLink,
    recamaras,
    edificio: building,
    tipologia,
    vista: view,
    asignacion: assignment,
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

    const planLink = normalizeViceroyPlanLink(rowDataByIndex(row, 1)); // B
    const building = String(rowDataByIndex(row, 2) || '').trim(); // C
    const tipologia = String(rowDataByIndex(row, 3) || '').trim(); // D
    const vista = String(rowDataByIndex(row, 4) || '').trim(); // E
    const m2 = rowDataByIndex(row, 14); // O
    const recRaw = rowDataByIndex(row, 15); // P
    const price = rowDataByIndex(row, 19); // T

    out.push({
      development: '',
      unidad,
      planLink,
      recamaras: normalizeViceroyRawBedroom(recRaw),
      edificio: building,
      tipologia,
      vista,
      asignacion: '',
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
  const currentScore = ['planLink', 'recamaras', 'edificio', 'tipologia', 'vista', 'asignacion', 'm2', 'sqft', 'price'].reduce((acc, key) => {
    return acc + (String(current[key] || '').trim() ? 1 : 0);
  }, 0);
  const nextScore = ['planLink', 'recamaras', 'edificio', 'tipologia', 'vista', 'asignacion', 'm2', 'sqft', 'price'].reduce((acc, key) => {
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
    const sourcePriority = floorDirs.indexOf(dir);
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
          entries.push({ name, fullPath, mtimeMs, sourcePriority });
        });
    } catch {}
  });

  const uniqueByName = new Map();
  entries.forEach((entry) => {
    const key = String(entry.name || '').toLowerCase();
    const current = uniqueByName.get(key);
    if (!current) {
      uniqueByName.set(key, entry);
      return;
    }
    const currentPriority = Number.isFinite(current.sourcePriority) ? current.sourcePriority : 999;
    const nextPriority = Number.isFinite(entry.sourcePriority) ? entry.sourcePriority : 999;
    if (nextPriority < currentPriority) {
      uniqueByName.set(key, entry);
      return;
    }
    if (nextPriority === currentPriority && entry.mtimeMs > current.mtimeMs) {
      uniqueByName.set(key, entry);
    }
  });
  const unique = Array.from(uniqueByName.values())
    .sort((a, b) => {
      const aPriority = Number.isFinite(a.sourcePriority) ? a.sourcePriority : 999;
      const bPriority = Number.isFinite(b.sourcePriority) ? b.sourcePriority : 999;
      if (aPriority !== bPriority) return aPriority - bPriority;
      if (a.mtimeMs !== b.mtimeMs) return b.mtimeMs - a.mtimeMs;
      return a.name.localeCompare(b.name);
    });

  function mappedNameStamp(name) {
    const match = String(name || '').match(/(\d{4})-(\d{2})-(\d{2})T(\d{2})-(\d{2})-(\d{2})-(\d{3})Z/i);
    if (!match) return 0;
    const iso = `${match[1]}-${match[2]}-${match[3]}T${match[4]}:${match[5]}:${match[6]}.${match[7]}Z`;
    const stamp = Date.parse(iso);
    return Number.isFinite(stamp) ? stamp : 0;
  }

  const floors = [];
  const loadedFiles = [];
  let filesToLoad = [];
  if (devSlug === 'viceroy-piloto') {
    const mapped = unique
      .filter((entry) => FLOOR_MAPPED_JSON_FILE_RE.test(entry.name))
      .sort((a, b) => {
        const aStamp = mappedNameStamp(a.name);
        const bStamp = mappedNameStamp(b.name);
        if (aStamp !== bStamp) return bStamp - aStamp;
        return b.mtimeMs - a.mtimeMs;
      });
    const canonical = unique
      .filter((entry) => FLOOR_JSON_FILE_RE.test(entry.name))
      .sort((a, b) => b.mtimeMs - a.mtimeMs);
    const rest = unique
      .filter((entry) => !FLOOR_MAPPED_JSON_FILE_RE.test(entry.name) && !FLOOR_JSON_FILE_RE.test(entry.name))
      .sort((a, b) => b.mtimeMs - a.mtimeMs);
    filesToLoad = [...mapped, ...canonical, ...rest];
  } else {
    filesToLoad = [...unique];
  }

  for (const entry of filesToLoad) {
    try {
      const raw = JSON.parse(fs.readFileSync(entry.fullPath, 'utf-8'));
      const rawDevSlug = String(raw && raw.developmentSlug || '').trim().toLowerCase();
      if (rawDevSlug && rawDevSlug !== devSlug) {
        continue;
      }
      const payloadFloors = Array.isArray(raw)
        ? raw
        : (raw && Array.isArray(raw.floors) ? raw.floors : (raw && raw.imageDataUrl ? [raw] : []));
      if (!payloadFloors.length) continue;
      // Guardrail: ignore clearly cross-loaded CEIBA payloads saved as "version-<other-dev>-...".
      if (devSlug !== DEFAULT_DEVELOPMENT_SLUG && /^version-.*\.json$/i.test(entry.name)) {
        const hasCeibaFloorNames = payloadFloors.some((floor) => String(floor && floor.name || '').toUpperCase().includes('CEIBA'));
        if (hasCeibaFloorNames) continue;
      }
      floors.push(...payloadFloors);
      loadedFiles.push(path.basename(entry.fullPath));
      if (devSlug === 'viceroy-piloto' && floors.length) break;
    } catch {}
  }
  return { floors, loadedFiles };
}

function readNamedFloorsByDevelopment(devSlug, requestedName) {
  const rawName = String(requestedName || '').trim();
  if (!rawName) return { floors: [], loadedFiles: [] };
  const safeName = sanitizeJsonFileName(rawName);
  if (!safeName) return { floors: [], loadedFiles: [] };
  const floorDirs = getDevelopmentFloorSearchDirs(devSlug);
  let filePath = '';
  for (const dir of floorDirs) {
    try {
      const files = fs.readdirSync(dir);
      const preferred = files.find((name) => String(name || '').toLowerCase() === safeName.toLowerCase());
      if (preferred) {
        filePath = path.join(dir, preferred);
        break;
      }
    } catch {}
  }
  if (!filePath || !fs.existsSync(filePath)) return { floors: [], loadedFiles: [] };
  try {
    const raw = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    const rawDevSlug = String(raw && raw.developmentSlug || '').trim().toLowerCase();
    if (rawDevSlug && rawDevSlug !== devSlug) return { floors: [], loadedFiles: [] };
    const payloadFloors = Array.isArray(raw)
      ? raw
      : (raw && Array.isArray(raw.floors) ? raw.floors : (raw && raw.imageDataUrl ? [raw] : []));
    if (!payloadFloors.length) return { floors: [], loadedFiles: [] };
    return { floors: payloadFloors, loadedFiles: [path.basename(filePath)] };
  } catch {
    return { floors: [], loadedFiles: [] };
  }
}

function persistSubmission(formatId, formatName, payload) {
  const current = readJson(SUBMISSIONS_PATH, []);
  const createdAt = new Date().toISOString();
  const payloadHash = crypto
    .createHash('sha1')
    .update(JSON.stringify(payload || {}))
    .digest('hex');
  const last = Array.isArray(current) && current.length ? current[current.length - 1] : null;
  if (last && last.formatId === formatId && last.payloadHash === payloadHash) {
    const lastTs = Date.parse(last.createdAt || '');
    const nowTs = Date.parse(createdAt);
    if (Number.isFinite(lastTs) && Number.isFinite(nowTs) && nowTs - lastTs <= 120000) {
      // Avoid duplicated writes when the same form is generated multiple times in short period.
      writeJson(DATA_PATH, payload);
      return last;
    }
  }
  const record = {
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`,
    createdAt,
    formatId,
    formatName,
    payloadHash,
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
  if (!merged.firmaConformidadLeyenda) {
    merged.firmaConformidadLeyenda = scope === 'NM'
      ? 'En representación del Titular'
      : 'Por su Propio Derecho';
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

function renderSimcaHome(req, res, options) {
  const currentEmail = String(req.user && req.user.email || '').toLowerCase();
  const isInternalUser = isInternalUserEmail(currentEmail);
  if (!isInternalUser) return res.redirect('/viceroy');
  const isGerente = currentEmail === GERENTE_EMAIL;
  const titleText = options && options.titleText ? String(options.titleText) : 'SIMCA';
  const subtitleText = options && options.subtitleText ? String(options.subtitleText) : 'Módulos internos de SIMCA.';
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
  const presentacionesCard = `
        <a class="card" href="/presentaciones">
          <span class="tag">Módulo</span>
          <h2 class="name">Presentaciones</h2>
          <p class="desc">Generación de presentaciones por proyecto con unidades seleccionadas.</p>
        </a>`;
  const tablaPagosEditorCard = isGerente ? `
        <a class="card" href="/tabla-pagos/editor">
          <span class="tag">Editor</span>
          <h2 class="name">Editor PDF Tabla de Pagos</h2>
          <p class="desc">Ajusta formato, columnas y layout de impresión del módulo Tabla de Pagos.</p>
        </a>` : '';
  const tablaPagosCard = `
        <a class="card" href="/tabla-pagos">
          <span class="tag">Módulo</span>
          <h2 class="name">Tabla de Pagos</h2>
          <p class="desc">Calcula enganche, pagos semestrales y balance de entrega por esquema.</p>
        </a>
        ${tablaPagosEditorCard}`;
  const brokersCard = isGerente ? `
        <a class="card" href="/brokers.simca.mx">
          <span class="tag">Módulo</span>
          <h2 class="name">Brokers.simca.mx</h2>
          <p class="desc">Editor visual de fichas HTML listo para copiar a Wix y exportar a PDF.</p>
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
  return res.send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>${escapeHtml(titleText)}</title>
  <style>
    :root{--bg:#f4f1e8;--card:#ffffff;--ink:#1a1a1a;--muted:#5f5f5f;--accent:#ffe816;--line:#d8d1c1;}
    *{box-sizing:border-box}
    body{font-family: Arial, sans-serif; margin:0; background:var(--bg); color:var(--ink);}
    .wrap{max-width:1100px; margin:0 auto; padding:32px 20px 60px;}
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
          <h1>${escapeHtml(titleText)}</h1>
          <p class="sub">${escapeHtml(subtitleText)}</p>
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
        ${brokersCard}
        ${presentacionesCard}
        ${tablaPagosCard}
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
}

app.get('/', requireAuth, (req, res) => {
  const currentEmail = String(req.user && req.user.email || '').toLowerCase();
  const isInternalUser = isInternalUserEmail(currentEmail);
  if (!isInternalUser) return res.redirect('/viceroy');
  return res.send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Backend Martin Barroso</title>
  <style>
    :root{--bg:#f4f1e8;--card:#ffffff;--ink:#1a1a1a;--muted:#5f5f5f;--accent:#ffe816;--line:#d8d1c1;}
    *{box-sizing:border-box}
    body{font-family:Arial,sans-serif;margin:0;background:var(--bg);color:var(--ink);}
    .wrap{max-width:980px;margin:0 auto;padding:32px 20px 60px;}
    .top{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;margin-bottom:26px;}
    h1{margin:0 0 6px;font-size:30px;}
    .sub{margin:0;color:var(--muted);}
    .user{font-size:13px;color:#5f5f5f;text-align:right;}
    .logout{display:inline-block;margin-top:8px;padding:8px 10px;border:1px solid #bdb8a9;border-radius:10px;background:#fff;color:#111;text-decoration:none;font-size:13px;font-weight:600;}
    .grid{display:grid;grid-template-columns:repeat(2,minmax(260px,1fr));gap:16px;}
    .card{display:block;background:var(--card);border:1px solid #dcd7cb;border-radius:18px;padding:22px;text-decoration:none;color:inherit;min-height:200px;}
    .card:hover{border-color:#b9b39f}
    .tag{display:inline-block;font-size:12px;font-weight:700;background:var(--accent);padding:4px 8px;border-radius:999px;margin-bottom:12px}
    .name{font-size:34px;margin:0 0 10px}
    .desc{margin:0;color:var(--muted);font-size:16px;line-height:1.45}
    @media (max-width:760px){.wrap{padding:20px 14px 36px}.grid{grid-template-columns:1fr}h1{font-size:24px}.name{font-size:28px}}
  </style></head><body>
    <div class="wrap">
      <div class="top">
        <div>
          <h1>Backend Martin Barroso</h1>
          <p class="sub">Selecciona un entorno para entrar a los módulos.</p>
        </div>
        <div class="user">
          <div>${escapeHtml(String(req.user && req.user.email || ''))}</div>
          <a class="logout" href="/logout">Cerrar sesión</a>
        </div>
      </div>
      <div class="grid">
        <a class="card" href="/simca">
          <span class="tag">Módulo</span>
          <h2 class="name">SIMCA</h2>
          <p class="desc">Acceso a FAES, ROI, Brokers, Presentaciones, Tabla de Pagos, Owner Services y herramientas comerciales.</p>
        </a>
        <a class="card" href="/viceroy">
          <span class="tag">Módulo</span>
          <h2 class="name">VICEROY</h2>
          <p class="desc">Acceso a Whisperlist, Registros, Tabla de Pago Viceroy, editor PDF y módulos internos de Viceroy.</p>
        </a>
      </div>
    </div>
  </body></html>`);
});

app.get('/simca', requireAuth, (req, res) => {
  return renderSimcaHome(req, res, {
    titleText: 'SIMCA',
    subtitleText: 'Módulos internos de SIMCA.'
  });
});

app.use('/legacy', requireInternalUser);
app.use('/generador-faes', requireInternalUser);
app.use('/plds', requireInternalUser);
app.use('/generador-roi', requireInternalUser);
app.use('/tabla-pagos', requireInternalUser);
app.use('/form', requireInternalUser);
app.use('/form-nacional', requireInternalUser);
app.use('/form-nacional-moral', requireInternalUser);
app.use('/form-extranjera-moral', requireInternalUser);
app.use('/format', requireInternalUser);
app.use('/format-nacional', requireInternalUser);
app.use('/format-nacional-moral', requireInternalUser);
app.use('/format-extranjera-moral', requireInternalUser);
app.use('/submissions', requireInternalUser);
app.use('/api/plds', requireInternalUser);
app.use('/api/roi', requireInternalUser);
app.use('/api/tabla-pagos', requireInternalUser);
app.use('/brokers.simca.mx', requireGerente);
app.use('/api/brokers-simca-mx', requireGerente);
app.use('/presentaciones', requireInternalUser);
app.use('/api/presentaciones', requireInternalUser);
app.use('/viceroy', requireAuth);
app.use('/viceroy/reservas', requireAuth);
app.use('/api/viceroy/reservas', requireAuth);
app.use('/viceroy/inicio', requireGerente);
app.use('/viceroy-piloto', requireGerente);
app.use('/api/viceroy-piloto', (req, res, next) => {
  return requireGerente(req, res, next);
});
app.use('/viceroy/registros', requireAuth);
app.use('/api/viceroy/registros', requireAuth);
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
  res.redirect('/form-extranjera-moral');
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

app.get('/api/session-info', requireAuth, (req, res) => {
  const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
  return res.json({
    ok: true,
    email: currentEmail,
    isGerente: currentEmail === GERENTE_EMAIL,
    isInternalUser: isInternalUserEmail(currentEmail)
  });
});

app.get('/tabla-pagos', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'tabla-pagos.html'));
});

app.get('/tabla-pagos/editor', requireGerente, (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'tabla-pagos-editor.html'));
});

app.get('/api/tabla-pagos/layout', (req, res) => {
  return res.json({ ok: true, layout: readTablaPagosLayout() });
});

app.post('/api/tabla-pagos/layout', requireGerente, (req, res) => {
  try {
    const incoming = req.body && req.body.layout ? req.body.layout : req.body;
    const saved = saveTablaPagosLayout(incoming);
    return res.json({ ok: true, layout: saved });
  } catch (err) {
    return res.status(500).json({
      error: 'No se pudo guardar el layout de Tabla de Pagos',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.post('/api/tabla-pagos/render-pdf', async (req, res) => {
  const html = typeof req.body?.html === 'string' ? req.body.html : '';
  const rawPrefix = String(req.body?.fileNamePrefix || 'tabla-pagos-simca').trim();
  const fileNamePrefix = rawPrefix.replace(/[^\w\- ]+/g, '').replace(/\s+/g, '-').slice(0, 60) || 'tabla-pagos-simca';
  if (!html || html.length < 100) {
    return res.status(400).json({ error: 'HTML inválido para generar PDF.' });
  }

  try {
    let pdfBuffer;
    try {
      const browser = await getSharedPdfBrowser();
      pdfBuffer = await buildPdfBufferWithBrowser(browser, html, {
        format: 'Letter',
        margin: { top: '0in', right: '0in', bottom: '0in', left: '0in' },
        printBackground: true
      });
    } catch (firstErr) {
      if (!isRetryablePdfError(firstErr)) throw firstErr;
      try {
        if (sharedPdfBrowser) await sharedPdfBrowser.close();
      } catch {}
      sharedPdfBrowser = null;
      const retryBrowser = await getSharedPdfBrowser();
      pdfBuffer = await buildPdfBufferWithBrowser(retryBrowser, html, {
        format: 'Letter',
        margin: { top: '0in', right: '0in', bottom: '0in', left: '0in' },
        printBackground: true
      });
    }

    const fileName = `${fileNamePrefix}.pdf`;
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    return res.send(pdfBuffer);
  } catch (err) {
    log(`Error en /api/tabla-pagos/render-pdf: ${err && err.stack ? err.stack : err}`);
    return res.status(500).json({
      error: 'No se pudo generar el PDF de Tabla de Pagos.',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/viceroy/tabla-pagos', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'tabla-pagos-viceroy.html'));
});

app.get('/viceroy/tabla-pagos/editor', requireGerente, (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'tabla-pagos-viceroy-editor.html'));
});

app.get('/api/viceroy/tabla-pagos/layout', (req, res) => {
  return res.json({ ok: true, layout: readViceroyPaymentPlanLayout() });
});

app.post('/api/viceroy/tabla-pagos/layout', requireGerente, (req, res) => {
  try {
    const incoming = req.body && req.body.layout ? req.body.layout : req.body;
    const saved = saveViceroyPaymentPlanLayout(incoming);
    return res.json({ ok: true, layout: saved });
  } catch (err) {
    return res.status(500).json({
      error: 'No se pudo guardar el layout de tabla de pagos Viceroy',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
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

app.get('/presentaciones', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'presentaciones.html'));
});

app.get('/brokers.simca.mx', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'brokers-simca-mx.html'));
});

app.get('/api/brokers-simca-mx/template', (req, res) => {
  return res.json({ ok: true, template: readBrokersSimcaTemplate() });
});

app.post('/api/brokers-simca-mx/template', (req, res) => {
  try {
    const incoming = req.body && req.body.template ? req.body.template : req.body;
    const saved = saveBrokersSimcaTemplate(incoming);
    return res.json({ ok: true, template: saved });
  } catch (err) {
    return res.status(500).json({
      error: 'No se pudo guardar la plantilla de Brokers.simca.mx',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/presentaciones/ceiba', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'presentaciones-ceiba.html'));
});

app.get('/presentaciones/solar-midtown', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'presentaciones-solar-midtown.html'));
});

app.get('/presentaciones/solar-midtown/editor', requireSolarMidtownEditor, (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'presentaciones-solar-midtown-editor.html'));
});

app.get('/api/presentaciones/solar-midtown/rows', (req, res) => {
  const data = readSolarMidtownRowsFromCbs();
  return res.json({
    ok: !data.error,
    sourceFile: data.sourceFile,
    totalRows: data.rows.length,
    rows: data.rows,
    crops: readSolarMidtownCropMap(),
    error: data.error || ''
  });
});

app.get('/api/presentaciones/ceiba/rows', (req, res) => {
  const data = readCeibaRowsFromCbs();
  return res.json({
    ok: !data.error,
    sourceFile: data.sourceFile,
    totalRows: data.rows.length,
    rows: data.rows,
    crops: readCeibaCropMap(),
    error: data.error || ''
  });
});

app.get('/api/presentaciones/ceiba/editor-access', (req, res) => {
  return res.json({
    ok: true,
    canEdit: hasCeibaEditorAccess(req),
    editorEmail: CEIBA_EDITOR_EMAIL
  });
});

app.get('/api/presentaciones/ceiba/crops', (req, res) => {
  return res.json({ ok: true, crops: readCeibaCropMap() });
});

app.post('/api/presentaciones/ceiba/crop', requireCeibaEditor, (req, res) => {
  try {
    const id = String(req.body && req.body.id || '').trim();
    if (!id) {
      return res.status(400).json({ error: 'Falta id de unidad.' });
    }
    const shouldReset = Boolean(req.body && req.body.reset);
    const map = readCeibaCropMap();
    if (shouldReset) {
      delete map[id];
      saveCeibaCropMap(map);
      return res.json({ ok: true, id, crop: defaultCeibaPlanCrop(), reset: true });
    }
    const crop = normalizeCeibaPlanCrop(req.body && req.body.crop);
    map[id] = crop;
    saveCeibaCropMap(map);
    return res.json({ ok: true, id, crop, reset: false });
  } catch (err) {
    return res.status(500).json({
      error: 'No se pudo guardar el recorte del plano.',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/api/presentaciones/solar-midtown/editor-access', (req, res) => {
  return res.json({
    ok: true,
    canEdit: hasSolarMidtownEditorAccess(req),
    editorEmail: SOLAR_MIDTOWN_EDITOR_EMAIL
  });
});

app.get('/api/presentaciones/solar-midtown/crops', (req, res) => {
  return res.json({ ok: true, crops: readSolarMidtownCropMap() });
});

app.post('/api/presentaciones/solar-midtown/crop', requireSolarMidtownEditor, (req, res) => {
  try {
    const id = String(req.body && req.body.id || '').trim();
    if (!id) {
      return res.status(400).json({ error: 'Falta id de unidad.' });
    }
    const shouldReset = Boolean(req.body && req.body.reset);
    const map = readSolarMidtownCropMap();
    if (shouldReset) {
      delete map[id];
      saveSolarMidtownCropMap(map);
      return res.json({ ok: true, id, crop: defaultSolarMidtownPlanCrop(), reset: true });
    }
    const crop = normalizeSolarMidtownPlanCrop(req.body && req.body.crop);
    map[id] = crop;
    saveSolarMidtownCropMap(map);
    return res.json({ ok: true, id, crop, reset: false });
  } catch (err) {
    return res.status(500).json({
      error: 'No se pudo guardar el recorte del plano.',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/api/presentaciones/solar-midtown/layout', (req, res) => {
  const layout = readSolarMidtownLayout();
  return res.json({ ok: true, layout });
});

app.get('/api/presentaciones/solar-midtown/presets-debug', (req, res) => {
  const repoCropsRaw = readJson(SOLAR_MIDTOWN_CROPS_REPO_PATH, {});
  const seedCropsRaw = readJson(SOLAR_MIDTOWN_CROPS_SEED_PATH, {});
  const runtimeCropsRaw = readJson(SOLAR_MIDTOWN_CROPS_PATH, {});
  const repoCrops = normalizeSolarMidtownCropMap(repoCropsRaw);
  const seedCrops = normalizeSolarMidtownCropMap(seedCropsRaw);
  const runtimeCrops = normalizeSolarMidtownCropMap(runtimeCropsRaw);
  const mergedCrops = readSolarMidtownCropMap();

  const repoLayoutRaw = readJson(SOLAR_MIDTOWN_LAYOUT_REPO_PATH, null);
  const seedLayoutRaw = readJson(SOLAR_MIDTOWN_LAYOUT_SEED_PATH, null);
  const runtimeLayoutRaw = readJson(SOLAR_MIDTOWN_LAYOUT_PATH, null);
  const mergedLayout = readSolarMidtownLayout();

  const data = readSolarMidtownRowsFromCbs();
  const rowIds = new Set((data.rows || []).map((row) => String(row && row.id || '').trim()).filter(Boolean));
  let cropCoverage = 0;
  rowIds.forEach((id) => {
    if (Object.prototype.hasOwnProperty.call(mergedCrops, id)) cropCoverage += 1;
  });

  return res.json({
    ok: true,
    paths: {
      repoCrops: SOLAR_MIDTOWN_CROPS_REPO_PATH,
      seedCrops: SOLAR_MIDTOWN_CROPS_SEED_PATH,
      runtimeCrops: SOLAR_MIDTOWN_CROPS_PATH,
      repoLayout: SOLAR_MIDTOWN_LAYOUT_REPO_PATH,
      seedLayout: SOLAR_MIDTOWN_LAYOUT_SEED_PATH,
      runtimeLayout: SOLAR_MIDTOWN_LAYOUT_PATH
    },
    exists: {
      repoCrops: fs.existsSync(SOLAR_MIDTOWN_CROPS_REPO_PATH),
      seedCrops: fs.existsSync(SOLAR_MIDTOWN_CROPS_SEED_PATH),
      runtimeCrops: fs.existsSync(SOLAR_MIDTOWN_CROPS_PATH),
      repoLayout: fs.existsSync(SOLAR_MIDTOWN_LAYOUT_REPO_PATH),
      seedLayout: fs.existsSync(SOLAR_MIDTOWN_LAYOUT_SEED_PATH),
      runtimeLayout: fs.existsSync(SOLAR_MIDTOWN_LAYOUT_PATH)
    },
    counts: {
      repoCrops: Object.keys(repoCrops).length,
      seedCrops: Object.keys(seedCrops).length,
      runtimeCrops: Object.keys(runtimeCrops).length,
      mergedCrops: Object.keys(mergedCrops).length,
      rows: rowIds.size,
      cropCoverage
    },
    hashes: {
      repoCropsFile: sha256ForFile(SOLAR_MIDTOWN_CROPS_REPO_PATH),
      seedCropsFile: sha256ForFile(SOLAR_MIDTOWN_CROPS_SEED_PATH),
      runtimeCropsFile: sha256ForFile(SOLAR_MIDTOWN_CROPS_PATH),
      mergedCrops: sha256ForJson(mergedCrops),
      repoLayoutFile: sha256ForFile(SOLAR_MIDTOWN_LAYOUT_REPO_PATH),
      seedLayoutFile: sha256ForFile(SOLAR_MIDTOWN_LAYOUT_SEED_PATH),
      runtimeLayoutFile: sha256ForFile(SOLAR_MIDTOWN_LAYOUT_PATH),
      mergedLayout: sha256ForJson(mergedLayout)
    },
    sample: {
      mergedCropUnit117: mergedCrops['solar-mt-unit-117'] || null,
      mergedLayout
    },
    layoutDiffersFromDefault: !isJsonEqual(mergedLayout, normalizeSolarMidtownLayout(defaultSolarMidtownLayout()))
  });
});

app.get('/api/presentaciones/solar-midtown/page-size', async (req, res) => {
  const brochurePath = getSolarMidtownBrochurePath(req.query && req.query.lang);
  const size = await readSolarMidtownBrochurePageSize(brochurePath);
  return res.json({ ok: true, pageSize: size });
});

app.get('/api/presentaciones/solar-midtown/brochure-debug', (req, res) => {
  const lang = normalizeSolarMidtownLang(req.query && req.query.lang);
  const candidates = getSolarMidtownBrochureCandidates(lang);
  const checks = candidates.map((p) => {
    const isUrl = /^https?:\/\//i.test(String(p || '').trim());
    let exists = false;
    let size = 0;
    if (isUrl) {
      exists = true;
    } else {
      try {
        if (fs.existsSync(p)) {
          exists = true;
          size = fs.statSync(p).size;
        }
      } catch {}
    }
    return { path: p, exists, size, sourceType: isUrl ? 'url' : 'file' };
  });
  return res.json({
    ok: true,
    lang,
    selectedPath: getSolarMidtownBrochurePath(lang),
    dataDir: DATA_DIR,
    repoDataDir: REPO_DATA_DIR,
    checks
  });
});

app.post('/api/presentaciones/solar-midtown/render-sample', requireSolarMidtownEditor, async (req, res) => {
  try {
    const incomingLayout = req.body && req.body.layout ? req.body.layout : req.body;
    const layout = normalizeSolarMidtownLayout(incomingLayout);
    const lang = normalizeSolarMidtownLang(req.body && req.body.lang);
    const sampleId = String(req.body && req.body.sampleId || '').trim();
    const data = readSolarMidtownRowsFromCbs();
    if (!Array.isArray(data.rows) || !data.rows.length) {
      return res.status(400).send('<!doctype html><html lang="es"><body>Sin datos de Solar Midtown.</body></html>');
    }
    const sampleRow = data.rows.find((row) => String(row && row.id || '') === sampleId) || data.rows[0];
    const brochurePath = getSolarMidtownBrochurePath(lang);
    const pageSize = await readSolarMidtownBrochurePageSize(brochurePath);
    const pagesHtml = buildSolarMidtownPagesHtml([sampleRow], layout, { absoluteAssets: true, lang });
    const html = buildSolarMidtownPagesDocument({
      title: 'Vista real · Solar Midtown',
      contentTop: '',
      pagesHtml,
      pdfMode: true,
      pageWidthPt: pageSize.widthPt,
      pageHeightPt: pageSize.heightPt,
      layout
    });
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    return res.send(html);
  } catch (err) {
    return res.status(500).send(`<!doctype html><html lang="es"><body>Error render sample: ${escapeHtml(err && err.message ? err.message : 'desconocido')}</body></html>`);
  }
});

app.post('/api/presentaciones/solar-midtown/layout', requireSolarMidtownEditor, (req, res) => {
  try {
    const incoming = req.body && req.body.layout ? req.body.layout : req.body;
    const saved = saveSolarMidtownLayout(incoming);
    return res.json({ ok: true, layout: saved });
  } catch (err) {
    return res.status(500).json({
      error: 'No se pudo guardar layout de Solar Midtown',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/api/presentaciones/solar-midtown/brochure', async (req, res) => {
  const brochurePath = getSolarMidtownBrochurePath(req.query && req.query.lang);
  const bytes = await readSolarMidtownBrochureBytes(brochurePath);
  if (!bytes) {
    return res.status(404).json({
      error: 'No se encontró brochure local de Solar Midtown.',
      expectedPath: brochurePath
    });
  }
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Cache-Control', 'private, max-age=300');
  return res.send(bytes);
});

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function parseSolarMidtownSelectedIds(rawIds, allRows) {
  const text = String(rawIds || '').trim();
  if (!text) return [];
  if (text === '__all' || text === '*') {
    const rows = Array.isArray(allRows) ? allRows : [];
    return rows.map((row) => String(row && row.id || '').trim()).filter(Boolean);
  }
  return text
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function normalizeSolarMidtownPdfQuality(raw) {
  const value = String(raw || '').trim().toLowerCase();
  if (value === 'high' || value === 'alta') return 'high';
  if (value === 'medium' || value === 'media') return 'medium';
  return 'low';
}

function getSolarMidtownPdfRenderProfile(qualityRaw) {
  const quality = normalizeSolarMidtownPdfQuality(qualityRaw);
  if (quality === 'high') {
    return { quality, imageFormat: 'jpeg', jpegQuality: 0.9, maxSide: 2600 };
  }
  if (quality === 'medium') {
    return { quality, imageFormat: 'jpeg', jpegQuality: 0.8, maxSide: 1800 };
  }
  return { quality: 'low', imageFormat: 'jpeg', jpegQuality: 0.68, maxSide: 1200 };
}

function parseSolarMidtownNumber(raw) {
  if (raw == null) return null;
  if (typeof raw === 'number' && Number.isFinite(raw)) return raw;
  const text = String(raw).replace(/,/g, '.');
  const match = text.match(/-?\d+(?:\.\d+)?/);
  if (!match) return null;
  const num = Number(match[0]);
  return Number.isFinite(num) ? num : null;
}

const SOLAR_MIDTOWN_VERTICAL_UNITS = new Set([
  101, 102, 103, 117, 118, 120, 220, 217, 201, 202,
  318, 317, 301, 302, 303, 401, 402, 403, 502
]);

function isSolarMidtownVerticalUnit(unidadRaw) {
  const num = parseSolarMidtownNumber(unidadRaw);
  if (!Number.isFinite(num)) return false;
  return SOLAR_MIDTOWN_VERTICAL_UNITS.has(Math.trunc(num));
}

function hasSolarMidtownDiscount(row) {
  const pct = Number(row && row.descuentoPct);
  const val = Number(row && row.descuentoValor);
  const list = Number(row && row.precioLista);
  const final = Number(row && row.precioFinal);
  if (Number.isFinite(pct) && Math.abs(pct) > 0.0001) return true;
  if (Number.isFinite(val) && Math.abs(val) > 0.0001) return true;
  if (Number.isFinite(list) && Number.isFinite(final) && Math.abs(list - final) > 0.01) return true;
  return false;
}

function parseSolarMidtownBrochurePages(raw, totalPages) {
  const total = Number(totalPages);
  if (!Number.isFinite(total) || total <= 0) return [];
  const safeTotal = Math.floor(total);
  const text = String(raw || '').trim();
  if (!text) {
    return Array.from({ length: safeTotal }, (_, idx) => idx);
  }
  const unique = new Set();
  text.split(',').forEach((piece) => {
    const part = String(piece || '').trim();
    if (!part) return;
    const rangeMatch = part.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
      const from = Math.max(1, Number(rangeMatch[1]));
      const to = Math.min(safeTotal, Number(rangeMatch[2]));
      if (Number.isFinite(from) && Number.isFinite(to)) {
        const start = Math.min(from, to);
        const end = Math.max(from, to);
        for (let page = start; page <= end; page += 1) unique.add(page - 1);
      }
      return;
    }
    const single = Number(part);
    if (!Number.isFinite(single)) return;
    const page = Math.floor(single);
    if (page < 1 || page > safeTotal) return;
    unique.add(page - 1);
  });
  return Array.from(unique).sort((a, b) => a - b);
}

function normalizeSolarMidtownLang(raw) {
  const value = String(raw || '').trim().toLowerCase();
  return value === 'en' ? 'en' : 'es';
}

function resolveFirstExistingPath(candidates) {
  const list = Array.isArray(candidates) ? candidates : [];
  for (const candidate of list) {
    const p = String(candidate || '').trim();
    if (!p) continue;
    if (/^https?:\/\//i.test(p)) return p;
    try {
      if (fs.existsSync(p)) return p;
    } catch {}
  }
  return String(list.find((item) => String(item || '').trim()) || '').trim();
}

function getSolarMidtownBrochurePath(langInput) {
  const lang = normalizeSolarMidtownLang(langInput);
  if (lang === 'en') {
    return resolveFirstExistingPath([
      SOLAR_MIDTOWN_BROCHURE_EN_ENV_URL,
      SOLAR_MIDTOWN_BROCHURE_EN_ENV_PATH,
      path.join(DATA_DIR, 'SOLAR Midtown ENG.pdf'),
      path.join(REPO_DATA_DIR, 'SOLAR Midtown ENG.pdf'),
      SOLAR_MIDTOWN_BROCHURE_ENG_PATH,
      SOLAR_MIDTOWN_BROCHURE_EN_FALLBACK_URL
    ]);
  }
  return resolveFirstExistingPath([
    SOLAR_MIDTOWN_BROCHURE_ES_ENV_URL,
    SOLAR_MIDTOWN_BROCHURE_ES_ENV_PATH,
    path.join(DATA_DIR, 'SOLAR Midtown ESP.pdf'),
    path.join(REPO_DATA_DIR, 'SOLAR Midtown ESP.pdf'),
    SOLAR_MIDTOWN_BROCHURE_PATH,
    SOLAR_MIDTOWN_BROCHURE_ES_FALLBACK_URL
  ]);
}

function getSolarMidtownBrochureCandidates(langInput) {
  const lang = normalizeSolarMidtownLang(langInput);
  if (lang === 'en') {
    return [
      SOLAR_MIDTOWN_BROCHURE_EN_ENV_URL,
      SOLAR_MIDTOWN_BROCHURE_EN_ENV_PATH,
      path.join(DATA_DIR, 'SOLAR Midtown ENG.pdf'),
      path.join(REPO_DATA_DIR, 'SOLAR Midtown ENG.pdf'),
      SOLAR_MIDTOWN_BROCHURE_ENG_PATH,
      SOLAR_MIDTOWN_BROCHURE_EN_FALLBACK_URL
    ].filter((p) => String(p || '').trim());
  }
  return [
    SOLAR_MIDTOWN_BROCHURE_ES_ENV_URL,
    SOLAR_MIDTOWN_BROCHURE_ES_ENV_PATH,
    path.join(DATA_DIR, 'SOLAR Midtown ESP.pdf'),
    path.join(REPO_DATA_DIR, 'SOLAR Midtown ESP.pdf'),
    SOLAR_MIDTOWN_BROCHURE_PATH,
    SOLAR_MIDTOWN_BROCHURE_ES_FALLBACK_URL
  ].filter((p) => String(p || '').trim());
}

function getSolarMidtownBrochureBaseName(sourcePathOrUrl, lang) {
  const fallback = lang === 'en' ? 'SOLAR Midtown ENG' : 'SOLAR Midtown ESP';
  const raw = String(sourcePathOrUrl || '').trim();
  if (!raw) return fallback;
  try {
    if (/^https?:\/\//i.test(raw)) {
      const u = new URL(raw);
      const name = path.basename(decodeURIComponent(u.pathname || ''));
      const base = String(name || '').replace(/\.pdf$/i, '').trim();
      return base || fallback;
    }
    const name = path.basename(raw);
    const base = String(name || '').replace(/\.pdf$/i, '').trim();
    return base || fallback;
  } catch {
    return fallback;
  }
}

function buildSolarMidtownPagesHtml(selectedRows, layoutInput, options) {
  const layout = normalizeSolarMidtownLayout(layoutInput);
  const opts = options && typeof options === 'object' ? options : {};
  const absoluteAssets = Boolean(opts.absoluteAssets);
  const lang = normalizeSolarMidtownLang(opts.lang);
  const t = lang === 'en'
    ? {
      titleFallback: 'Option',
      development: 'Development',
      department: 'Condo',
      sqm: 'SQFT',
      priceSummary: 'Price summary',
      listPrice: 'List price',
      discountPct: 'Discount %',
      discountValue: 'Discount amount',
      finalPrice: 'Final price',
      paymentPlan: 'Payment plan',
      downPayment: '30% Down payment',
      months1: '10% at 6 months (1)',
      months2: '10% at 6 months (2)',
      delivery: '50% On delivery'
    }
    : {
      titleFallback: 'Página adicional',
      development: 'Desarrollo',
      department: 'Departamento',
      sqm: 'Metros cuadrados',
      priceSummary: 'Resumen de precio',
      listPrice: 'Precio lista',
      discountPct: 'Descuento %',
      discountValue: 'Descuento valor',
      finalPrice: 'Precio final',
      paymentPlan: 'Forma de pago',
      downPayment: '30% Enganche',
      months1: '10% a 6 meses (1)',
      months2: '10% a 6 meses (2)',
      delivery: '50% Entrega'
    };
  const planDataUrlById = opts.planDataUrlById && typeof opts.planDataUrlById === 'object'
    ? opts.planDataUrlById
    : {};
  if (!selectedRows.length) {
    return '<section class="empty">No se seleccionaron unidades para la vista previa.</section>';
  }
  return selectedRows.map((row, index) => {
    const rowPlanDataUrl = planDataUrlById[String(row && row.id || '')] || '';
    const planPath = row.planLink
      ? `/api/presentaciones/solar-midtown/plan-image?url=${encodeURIComponent(row.planLink)}`
      : '';
    const planProxy = planPath
      ? (absoluteAssets ? `${APP_BASE_URL_NORMALIZED}${planPath}` : planPath)
      : '';
    const planSrc = rowPlanDataUrl || planProxy;
    const cropDataAttrs = buildSolarMidtownCropDataAttrs(row && row.crop, row, layout);
    const planBlock = planSrc
      ? `<img class="plan-image" ${cropDataAttrs} src="${escapeHtml(planSrc)}" alt="Plano ${escapeHtml(row.unidad || `Unidad ${index + 1}`)}" />`
      : '<div class="empty">Sin link de plano (columna B)</div>';
    const titleText = layout.title.text || t.titleFallback;
    const metros = String(row && row.colF || '').trim();
    const sqft = String(row && row.colG || '').trim();
    const discountApplies = hasSolarMidtownDiscount(row);
    return `<article class="unit-page">
      <header>
        ${layout.title.show ? `<h2 style="color:${escapeHtml(layout.title.color)};font-size:${layout.title.fontSize}px;font-weight:${layout.title.fontWeight};">${escapeHtml(titleText)} ${index + 1}</h2>` : ''}
        ${layout.subtitle.show ? `<p style="color:${escapeHtml(layout.subtitle.color)};font-size:${layout.subtitle.fontSize}px;">${escapeHtml(t.development)}: ${escapeHtml(row.desarrollo || 'SOLAR MT')}</p>` : ''}
      </header>
      <div class="unit-grid" style="display:grid;grid-template-columns:${layout.grid.leftRatio}fr ${layout.grid.rightRatio}fr;gap:${layout.grid.gap}px;">
        <section class="plan-box" style="background:${escapeHtml(layout.planBox.backgroundColor)};border:1px solid ${escapeHtml(layout.planBox.borderColor)};border-radius:${layout.planBox.borderRadius}px;min-height:${layout.planBox.minHeight}px;">${planBlock}</section>
        <section class="data-box" data-align="${escapeHtml(layout.priceBox.verticalAlign)}" style="background:${escapeHtml(layout.priceBox.backgroundColor)};border:1px solid ${escapeHtml(layout.priceBox.borderColor)};border-radius:${layout.priceBox.borderRadius}px;color:${escapeHtml(layout.priceBox.textColor)};">
          <div class="unit-meta-top">
            <p class="meta department-meta"><span class="meta-label">${escapeHtml(t.department)}:</span> <strong class="meta-value">${escapeHtml(row.unidad || '-')}</strong></p>
            <p class="meta meters-meta"><span class="meta-label">${escapeHtml(t.sqm)}:</span> <strong class="meta-value">${escapeHtml(lang === 'en' ? (sqft || '-') : (metros || '-'))}${lang === 'en' ? '' : ' m²'}</strong></p>
          </div>
          <h3>${escapeHtml(t.priceSummary)}</h3>
          <ul>
            <li><span>${escapeHtml(t.listPrice)}</span><strong>${escapeHtml(row.precioListaFmt || '-')}</strong></li>
            ${discountApplies ? `<li><span>${escapeHtml(t.discountPct)}</span><strong>${escapeHtml(row.descuentoPctFmt || '-')}</strong></li>` : ''}
            ${discountApplies ? `<li><span>${escapeHtml(t.discountValue)}</span><strong>${escapeHtml(row.descuentoValorFmt || '-')}</strong></li>` : ''}
            ${discountApplies ? `<li><span>${escapeHtml(t.finalPrice)}</span><strong>${escapeHtml(row.precioFinalFmt || '-')}</strong></li>` : ''}
          </ul>
          <h3>${escapeHtml(t.paymentPlan)}</h3>
          <ul>
            <li><span>${escapeHtml(t.downPayment)}</span><strong>${escapeHtml(row.pagos && row.pagos.enganche30Fmt || '-')}</strong></li>
            <li><span>${escapeHtml(t.months1)}</span><strong>${escapeHtml(row.pagos && row.pagos.seisMeses1Fmt || '-')}</strong></li>
            <li><span>${escapeHtml(t.months2)}</span><strong>${escapeHtml(row.pagos && row.pagos.seisMeses2Fmt || '-')}</strong></li>
            <li><span>${escapeHtml(t.delivery)}</span><strong>${escapeHtml(row.pagos && row.pagos.entrega50Fmt || '-')}</strong></li>
          </ul>
        </section>
      </div>
    </article>`;
  }).join('');
}

async function fetchImageAsDataUrl(url) {
  try {
    const parsed = new URL(String(url || '').trim());
    if (!['http:', 'https:'].includes(parsed.protocol)) return '';
    const upstream = await fetch(parsed.toString());
    if (!upstream.ok) return '';
    const contentType = upstream.headers.get('content-type') || 'image/png';
    const buf = Buffer.from(await upstream.arrayBuffer());
    if (!buf.length) return '';
    return `data:${contentType};base64,${buf.toString('base64')}`;
  } catch {
    return '';
  }
}

// PDF now reuses the exact same HTML renderer as preview to avoid crop mismatches.

async function readSolarMidtownBrochureBytes(sourceInput) {
  const source = String(sourceInput || '').trim();
  if (!source) return null;
  try {
    if (/^https?:\/\//i.test(source)) {
      const upstream = await fetch(source);
      if (!upstream.ok) return null;
      const bytes = Buffer.from(await upstream.arrayBuffer());
      return bytes.length ? bytes : null;
    }
    if (!fs.existsSync(source)) return null;
    const bytes = fs.readFileSync(source);
    return bytes.length ? bytes : null;
  } catch {
    return null;
  }
}

async function readSolarMidtownBrochurePageSize(brochurePathInput) {
  const fallback = { widthPt: 792, heightPt: 612 };
  const brochurePath = String(brochurePathInput || SOLAR_MIDTOWN_BROCHURE_PATH);
  try {
    const bytes = await readSolarMidtownBrochureBytes(brochurePath);
    if (!bytes) return fallback;
    const doc = await PDFDocument.load(bytes);
    const first = doc.getPage(0);
    if (!first) return fallback;
    const widthPt = first.getWidth();
    const heightPt = first.getHeight();
    if (!Number.isFinite(widthPt) || !Number.isFinite(heightPt) || widthPt <= 0 || heightPt <= 0) {
      return fallback;
    }
    return { widthPt, heightPt };
  } catch {
    return fallback;
  }
}

async function readSolarMidtownBrochureMeta(brochurePathInput) {
  const fallback = { widthPt: 792, heightPt: 612, pageCount: 0 };
  const brochurePath = String(brochurePathInput || SOLAR_MIDTOWN_BROCHURE_PATH);
  try {
    const bytes = await readSolarMidtownBrochureBytes(brochurePath);
    if (!bytes) return fallback;
    const doc = await PDFDocument.load(bytes);
    const pageCount = doc.getPageCount();
    const first = pageCount > 0 ? doc.getPage(0) : null;
    if (!first) return { ...fallback, pageCount };
    const widthPt = first.getWidth();
    const heightPt = first.getHeight();
    if (!Number.isFinite(widthPt) || !Number.isFinite(heightPt) || widthPt <= 0 || heightPt <= 0) {
      return { ...fallback, pageCount };
    }
    return { widthPt, heightPt, pageCount };
  } catch {
    return fallback;
  }
}

function buildSolarMidtownCropScript(options) {
  const opts = options && typeof options === 'object' ? options : {};
  const imageFormat = String(opts.imageFormat || 'png').toLowerCase() === 'jpeg' ? 'jpeg' : 'png';
  const jpegQuality = clampNumber(opts.jpegQuality, 0.82, 0.4, 0.98);
  const maxSide = clampNumber(opts.maxSide, 1800, 400, 5000);
  const qualityLiteral = Number(jpegQuality.toFixed(3));
  const maxSideLiteral = Math.round(maxSide);
  return `<script>
    (function () {
      function clamp(v, min, max) { return Math.max(min, Math.min(max, v)); }
      const OUT_FORMAT = '${imageFormat}';
      const JPEG_QUALITY = ${qualityLiteral};
      const MAX_SIDE = ${maxSideLiteral};
      function numAttr(el, name, fallback) {
        const value = Number(el.getAttribute(name));
        return Number.isFinite(value) ? value : fallback;
      }
      function cropPlanImage(img) {
        return new Promise(function (resolve) {
          const src = img.getAttribute('src');
          if (!src) return resolve();
          const helper = new Image();
          helper.crossOrigin = 'anonymous';
          helper.onload = function () {
            try {
              const iw = helper.naturalWidth || helper.width;
              const ih = helper.naturalHeight || helper.height;
              if (!iw || !ih) return resolve();
              const cx = numAttr(img, 'data-crop-x', 0);
              const cy = numAttr(img, 'data-crop-y', 0);
              const cw = clamp(numAttr(img, 'data-crop-w', 100), 1, 100);
              const ch = clamp(numAttr(img, 'data-crop-h', 100), 1, 100);
              const expand = clamp(numAttr(img, 'data-crop-expand', 1), 1, 2.4);
              const centerX = cx + (cw / 2);
              const centerY = cy + (ch / 2);
              const ew = clamp(cw * expand, 1, 100);
              const eh = clamp(ch * expand, 1, 100);
              const ex = centerX - (ew / 2);
              const ey = centerY - (eh / 2);
              const sx = Math.round((ex / 100) * iw);
              const sy = Math.round((ey / 100) * ih);
              const sw = Math.max(1, Math.round((ew / 100) * iw));
              const sh = Math.max(1, Math.round((eh / 100) * ih));
              const out = document.createElement('canvas');
              let outW = Math.max(1, sw);
              let outH = Math.max(1, sh);
              const longest = Math.max(outW, outH);
              if (longest > MAX_SIDE) {
                const ratio = MAX_SIDE / longest;
                outW = Math.max(1, Math.round(outW * ratio));
                outH = Math.max(1, Math.round(outH * ratio));
              }
              out.width = outW;
              out.height = outH;
              const outCtx = out.getContext('2d');
              if (!outCtx) return resolve();
              const srcX = Math.max(0, sx);
              const srcY = Math.max(0, sy);
              const srcX2 = Math.min(iw, sx + sw);
              const srcY2 = Math.min(ih, sy + sh);
              const srcW = Math.max(0, srcX2 - srcX);
              const srcH = Math.max(0, srcY2 - srcY);
              if (srcW > 0 && srcH > 0) {
                const destX = Math.max(0, -sx);
                const destY = Math.max(0, -sy);
                outCtx.drawImage(helper, srcX, srcY, srcW, srcH, Math.round(destX * (outW / sw)), Math.round(destY * (outH / sh)), Math.round(srcW * (outW / sw)), Math.round(srcH * (outH / sh)));
              }
              const mime = OUT_FORMAT === 'jpeg' ? 'image/jpeg' : 'image/png';
              img.src = OUT_FORMAT === 'jpeg'
                ? out.toDataURL(mime, JPEG_QUALITY)
                : out.toDataURL(mime);
            } catch {}
            resolve();
          };
          helper.onerror = function () { resolve(); };
          helper.src = src;
        });
      }
      window.__pdfReady = false;
      const images = Array.from(document.querySelectorAll('img.plan-image'));
      if (!images.length) {
        window.__pdfReady = true;
        return;
      }
      Promise.all(images.map(cropPlanImage)).finally(function () {
        window.__pdfReady = true;
      });
    })();
  </script>`;
}

function buildSolarMidtownPagesDocument(options) {
  const layout = normalizeSolarMidtownLayout(options && options.layout);
  const title = options && options.title ? options.title : 'Vista previa · Solar Midtown';
  const pagesHtml = options && options.pagesHtml ? options.pagesHtml : '';
  const contentTop = options && options.contentTop ? options.contentTop : '';
  const pdfMode = Boolean(options && options.pdfMode);
  const pageWidthPt = Number(options && options.pageWidthPt) > 0 ? Number(options.pageWidthPt) : 792;
  const pageHeightPt = Number(options && options.pageHeightPt) > 0 ? Number(options.pageHeightPt) : 612;
  const pageSizeCss = `${pageWidthPt}pt ${pageHeightPt}pt`;
  const pageBg = pdfMode ? layout.page.backgroundColor : '#f4f1e8';
  const wrapBg = 'transparent';
  const pageCss = pdfMode
    ? `@page { size: ${pageSizeCss}; margin: 0; }`
    : '';
  const wrapPadding = clampNumber(layout.page.padding, 16, 0, 80);
  const wrapCss = pdfMode
    ? `max-width:none;margin:0;padding:0;`
    : 'max-width:1160px;margin:0 auto;padding:22px 16px 40px;';
  const unitPageCss = pdfMode
    ? `background:#fff;border:0;border-radius:0;padding:${wrapPadding}pt;height:${Math.max(120, pageHeightPt)}pt;width:${Math.max(120, pageWidthPt)}pt;margin:0;page-break-inside:avoid;page-break-after:always;overflow:hidden;display:flex;flex-direction:column;`
    : 'background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px;margin-bottom:14px;page-break-inside:avoid;page-break-after:always;';
  const pageContainerCss = pdfMode
    ? `padding:0;margin:0;`
    : '';
  const renderProfile = options && options.renderProfile && typeof options.renderProfile === 'object'
    ? options.renderProfile
    : { imageFormat: 'png', jpegQuality: 0.9, maxSide: 2600 };
  return `<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>${escapeHtml(title)}</title>
  <style>
    ${pageCss}
    :root{
      --bg:${pageBg};
      --card:#fff;
      --line:#d9d3c3;
      --ink:#1f2330;
      --muted:#5f6572;
      --price-heading:${escapeHtml(layout.priceBox.headingColor)};
      --price-text:${escapeHtml(layout.priceBox.textColor)};
      --price-divider:${escapeHtml(layout.priceBox.dividerColor)};
      --price-heading-size:${layout.priceBox.headingSize}px;
      --price-body-size:${layout.priceBox.bodySize}px;
      --price-label-size:${layout.priceBox.labelSize}px;
      --price-value-size:${layout.priceBox.valueSize}px;
      --price-meta-size:${layout.priceBox.metaSize}px;
      --price-department-label-size:${layout.priceBox.departmentLabelSize || layout.priceBox.metaSize}px;
      --price-department-value-size:${layout.priceBox.departmentValueSize || layout.priceBox.metaSize}px;
      --price-meters-label-size:${layout.priceBox.metersLabelSize || layout.priceBox.metaSize}px;
      --price-meters-value-size:${layout.priceBox.metersValueSize || layout.priceBox.metaSize}px;
      --price-row-padding-y:${layout.priceBox.rowPaddingY}px;
      --price-line-height:${layout.priceBox.lineHeight};
      --price-word-spacing:${layout.priceBox.wordSpacing}px;
      --price-letter-spacing:${layout.priceBox.letterSpacing}px;
      --price-meta-top-offset:${layout.priceBox.metaTopOffset}px;
      --price-meta-gap:${layout.priceBox.metaGroupGap}px;
      --price-justify:${layout.priceBox.verticalAlign};
    }
    *{box-sizing:border-box}
    body{margin:0;font-family:Arial,sans-serif;background:var(--bg);color:var(--ink);}
    .wrap{${wrapCss}background:${wrapBg};}
    .head{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:14px;}
    h1{margin:0;font-size:28px;}
    .sub{margin:6px 0 0;color:var(--muted);}
    .btn{display:inline-block;padding:9px 12px;border:1px solid #beb8a8;border-radius:10px;background:#fff;color:#111;text-decoration:none;font-size:13px;font-weight:700;}
    .panel{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:14px;margin-bottom:14px;}
    iframe{width:100%;height:72vh;border:1px solid #d6d8df;border-radius:12px;background:#fff;}
    .unit-page{${unitPageCss}}
    .unit-page:last-child{page-break-after:auto;}
    .unit-page h2{margin:0 0 4px;font-size:22px;}
    .unit-page header p{margin:0;color:var(--muted);}
    .page-content{${pageContainerCss}}
    .unit-grid{display:grid;grid-template-columns:1.2fr 1fr;gap:14px;margin-top:10px;flex:1;min-height:0;}
    .plan-box{border:1px solid #d4d7df;border-radius:12px;background:#f8f8f8;display:grid;place-items:center;min-height:340px;overflow:hidden;}
    .unit-page .plan-box,.unit-page .data-box{height:100%;}
    .plan-box img{max-width:100%;max-height:100%;display:block;object-fit:contain;}
    .data-box{
      border:1px solid #df7d24;border-radius:12px;padding:14px;background:#f28b2c;color:#fff;
      display:flex;flex-direction:column;justify-content:flex-start;
    }
    .data-box[data-align="center"]{justify-content:center;}
    .data-box[data-align="end"]{justify-content:flex-end;}
    .data-box h3{
      margin:0 0 8px;font-size:var(--price-heading-size);color:var(--price-heading);
      line-height:var(--price-line-height);word-spacing:var(--price-word-spacing);letter-spacing:var(--price-letter-spacing);
    }
    .data-box ul{list-style:none;padding:0;margin:0 0 12px;}
    .data-box li{
      display:flex;justify-content:space-between;gap:12px;
      padding:var(--price-row-padding-y) 0;border-bottom:1px solid var(--price-divider);
      line-height:var(--price-line-height);word-spacing:var(--price-word-spacing);letter-spacing:var(--price-letter-spacing);
      font-size:var(--price-body-size);
    }
    .data-box li:last-child{border-bottom:0;}
    .data-box span{
      color:var(--price-text);font-size:var(--price-label-size);
      line-height:var(--price-line-height);word-spacing:var(--price-word-spacing);letter-spacing:var(--price-letter-spacing);
    }
    .data-box strong{
      color:var(--price-text);font-size:var(--price-value-size);
      line-height:var(--price-line-height);word-spacing:var(--price-word-spacing);letter-spacing:var(--price-letter-spacing);
    }
    .meta{
      margin:8px 0 0;color:var(--price-text);font-size:var(--price-meta-size);
      line-height:var(--price-line-height);word-spacing:var(--price-word-spacing);letter-spacing:var(--price-letter-spacing);
    }
    .unit-meta-top{margin-top:var(--price-meta-top-offset);margin-bottom:var(--price-meta-gap);}
    .unit-meta-top .meta{margin:0 0 4px;}
    .unit-meta-top .meta:last-child{margin-bottom:0;}
    .unit-meta-top .department-meta .meta-label{font-size:var(--price-department-label-size);}
    .unit-meta-top .department-meta .meta-value{font-size:var(--price-department-value-size);}
    .unit-meta-top .meters-meta .meta-label{font-size:var(--price-meters-label-size);}
    .unit-meta-top .meters-meta .meta-value{font-size:var(--price-meters-value-size);}
    .empty{display:grid;place-items:center;text-align:center;min-height:170px;border:1px dashed #cfd4df;border-radius:12px;color:var(--muted);padding:12px;background:#fff;}
    code{background:#f2f4f8;padding:2px 4px;border-radius:6px}
    @media (max-width:900px){.unit-grid{grid-template-columns:1fr;} iframe{height:58vh;}}
  </style></head><body>
    <div class="wrap">${contentTop}${pagesHtml}</div>
    ${buildSolarMidtownCropScript(renderProfile)}
  </body></html>`;
}

app.get('/api/presentaciones/solar-midtown/plan-image', async (req, res) => {
  try {
    const source = String(req.query && req.query.url || '').trim();
    if (!source) return res.status(400).json({ error: 'Falta parámetro url' });
    let parsed;
    try {
      parsed = new URL(source);
    } catch {
      return res.status(400).json({ error: 'URL inválida' });
    }
    if (!['http:', 'https:'].includes(parsed.protocol)) {
      return res.status(400).json({ error: 'Solo se permiten URLs http/https' });
    }
    const upstream = await fetch(parsed.toString());
    if (!upstream.ok) {
      return res.status(502).json({ error: `No se pudo descargar plano (${upstream.status})` });
    }
    const contentType = upstream.headers.get('content-type') || 'image/png';
    const data = Buffer.from(await upstream.arrayBuffer());
    res.setHeader('Content-Type', contentType);
    res.setHeader('Cache-Control', 'private, max-age=300');
    res.setHeader('Access-Control-Allow-Origin', '*');
    return res.send(data);
  } catch (err) {
    return res.status(500).json({
      error: 'Error al obtener la imagen del plano',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/api/presentaciones/ceiba/plan-image', async (req, res) => {
  try {
    const source = String(req.query && req.query.url || '').trim();
    if (!source) return res.status(400).json({ error: 'Falta parámetro url' });
    let parsed;
    try {
      parsed = new URL(source);
    } catch {
      return res.status(400).json({ error: 'URL inválida' });
    }
    if (!['http:', 'https:'].includes(parsed.protocol)) {
      return res.status(400).json({ error: 'Solo se permiten URLs http/https' });
    }
    const upstream = await fetch(parsed.toString());
    if (!upstream.ok) {
      return res.status(502).json({ error: `No se pudo descargar plano (${upstream.status})` });
    }
    const contentType = upstream.headers.get('content-type') || 'image/png';
    const data = Buffer.from(await upstream.arrayBuffer());
    res.setHeader('Content-Type', contentType);
    res.setHeader('Cache-Control', 'private, max-age=300');
    res.setHeader('Access-Control-Allow-Origin', '*');
    return res.send(data);
  } catch (err) {
    return res.status(500).json({
      error: 'Error al obtener la imagen del plano',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/presentaciones/solar-midtown/preview', async (req, res) => {
  const data = readSolarMidtownRowsFromCbs();
  const selectedIds = parseSolarMidtownSelectedIds(req.query && req.query.ids, data.rows);
  const renderProfile = getSolarMidtownPdfRenderProfile('low');
  const layout = readSolarMidtownLayout();
  const selectedRows = data.rows.filter((row) => selectedIds.includes(row.id));
  const brochurePathEs = getSolarMidtownBrochurePath('es');
  const brochureExists = Boolean(await readSolarMidtownBrochureBytes(brochurePathEs));
  const brochureMeta = brochureExists ? await readSolarMidtownBrochureMeta(brochurePathEs) : { pageCount: 0 };
  const selectedBrochurePages = parseSolarMidtownBrochurePages(req.query && req.query.pages, brochureMeta.pageCount);
  const selectedBrochurePagesCsv = selectedBrochurePages.map((idx) => String(idx + 1)).join(',');
  const idsRaw = selectedIds.join(',');
  const brochureEmbed = brochureExists
    ? `<iframe src="/api/presentaciones/solar-midtown/brochure?lang=es" title="Brochure Solar Midtown"></iframe>`
    : `<div class="empty">No se encontró brochure local en:<br><code>${escapeHtml(brochurePathEs)}</code></div>`;
  const brochurePagesPanel = brochureExists && brochureMeta.pageCount > 0 ? `<section class="panel">
      <h2 style="margin:0 0 8px;font-size:20px;">Hojas del brochure para descargar</h2>
      <p class="sub" style="margin:0 0 10px;">Marca las hojas que quieras conservar en el PDF final.</p>
      <div id="brochurePagesBox" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(82px,1fr));gap:8px;max-width:900px;">
        ${Array.from({ length: brochureMeta.pageCount }, (_, idx) => {
          const page = idx + 1;
          const checked = selectedBrochurePages.includes(idx) ? 'checked' : '';
          return `<label style="display:flex;align-items:center;gap:6px;border:1px solid #d5d9e3;padding:6px 8px;border-radius:8px;background:#fff;">
              <input type="checkbox" class="brochure-page-check" value="${page}" ${checked}>
              <span style="font-size:12px;">Hoja ${page}</span>
            </label>`;
        }).join('')}
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:10px;">
        <button type="button" class="btn" id="brochureSelectAllBtn">Seleccionar todo</button>
        <button type="button" class="btn" id="brochureSelectNoneBtn">Quitar todo</button>
      </div>
      <p id="brochurePagesHint" class="sub" style="margin:8px 0 0;">Seleccionadas: ${selectedBrochurePages.length} de ${brochureMeta.pageCount}</p>
    </section>` : '';
  const pagesHtml = `<section><h2 style="margin:0 0 8px;font-size:20px;">Páginas agregadas al final</h2>${buildSolarMidtownPagesHtml(selectedRows, layout, { lang: 'es' })}</section>`;
  const sampleIdForEditor = selectedRows[0] && selectedRows[0].id ? String(selectedRows[0].id) : '';
  const editorHref = `/presentaciones/solar-midtown/editor?ids=${encodeURIComponent(idsRaw)}${sampleIdForEditor ? `&sample=${encodeURIComponent(sampleIdForEditor)}` : ''}`;
  const contentTop = `<div class="head">
      <div>
        <h1>Vista previa · Solar Midtown</h1>
        <p class="sub">Brochure base + páginas adicionales por unidad seleccionada.</p>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;">
        <a class="btn" href="${editorHref}">Editar plantilla</a>
        <a class="btn" id="downloadWithPagesBtnEs" href="/api/presentaciones/solar-midtown/download.pdf?lang=es&quality=low&ids=${encodeURIComponent(idsRaw)}&pages=${encodeURIComponent(selectedBrochurePagesCsv)}">Descargar presentación (PDF ES)</a>
        <a class="btn" id="downloadWithPagesBtnEn" href="/api/presentaciones/solar-midtown/download.pdf?lang=en&quality=low&ids=${encodeURIComponent(idsRaw)}&pages=${encodeURIComponent(selectedBrochurePagesCsv)}">Descargar presentación (PDF EN)</a>
        <a class="btn" href="/presentaciones/solar-midtown">Volver</a>
      </div>
    </div>
    <div id="previewDownloadModal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.42);z-index:70;padding:18px;align-items:center;justify-content:center;">
      <section style="background:#fff;border:1px solid #d4d7df;border-radius:14px;max-width:560px;width:100%;padding:18px;">
        <h3 style="margin:0 0 6px;font-size:22px;color:#1f2330;"><span style="width:22px;height:22px;border-radius:50%;border:3px solid #dfe4ee;border-top-color:#1f8f3a;display:inline-block;vertical-align:middle;margin-right:8px;animation:spinPreviewDownload .8s linear infinite;"></span>Preparando descarga</h3>
        <p id="previewDownloadModalText" style="margin:0;color:#5f6572;line-height:1.45;">Por favor espera, estamos recolectando datos para generar tu presentación PDF.</p>
      </section>
    </div>
    <style>
      @keyframes spinPreviewDownload { to { transform: rotate(360deg); } }
    </style>
    ${brochurePagesPanel}
    <section class="panel">
      <h2 style="margin:0 0 8px;font-size:20px;">Brochure base</h2>
      ${brochureEmbed}
    </section>
    <script>
      (function () {
        const checks = Array.from(document.querySelectorAll('.brochure-page-check'));
        const hint = document.getElementById('brochurePagesHint');
        const downloadBtnEs = document.getElementById('downloadWithPagesBtnEs');
        const downloadBtnEn = document.getElementById('downloadWithPagesBtnEn');
        const selectAllBtn = document.getElementById('brochureSelectAllBtn');
        const selectNoneBtn = document.getElementById('brochureSelectNoneBtn');
        const modal = document.getElementById('previewDownloadModal');
        const modalText = document.getElementById('previewDownloadModalText');
        const selectedIdsList = ${JSON.stringify(selectedIds)};
        const CHUNK_SIZE = 24;
        if (!checks.length || !downloadBtnEs || !downloadBtnEn) return;
        const baseUrlEs = '/api/presentaciones/solar-midtown/download.pdf?lang=es&quality=low';
        const baseUrlEn = '/api/presentaciones/solar-midtown/download.pdf?lang=en&quality=low';
        function showModal(message) {
          if (!modal) return;
          if (modalText) modalText.textContent = message || 'Por favor espera, estamos recolectando datos para generar tu presentación PDF.';
          modal.style.display = 'flex';
        }
        function hideModal() {
          if (!modal) return;
          modal.style.display = 'none';
        }
        function parseFileName(disposition, fallback) {
          const raw = String(disposition || '');
          const m = raw.match(/filename\\*=UTF-8''([^;]+)|filename=\\"([^\\"]+)\\"|filename=([^;]+)/i);
          const encoded = m && (m[1] || m[2] || m[3]);
          if (!encoded) return fallback;
          try { return decodeURIComponent(String(encoded).replace(/^['"]|['"]$/g, '')); } catch { return String(encoded).replace(/^['"]|['"]$/g, ''); }
        }
        function splitIds(list, size) {
          const out = [];
          for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
          return out;
        }
        async function downloadCurrent(lang) {
          const pages = selectedPages();
          if (!pages.length) {
            alert('Selecciona al menos 1 hoja del brochure.');
            return;
          }
          const baseUrl = lang === 'en' ? baseUrlEn : baseUrlEs;
          const idChunks = splitIds(selectedIdsList, CHUNK_SIZE);
          const multi = idChunks.length > 1;
          const msg = lang === 'en'
            ? 'Please wait, we are collecting data to generate the English PDF presentation.'
            : 'Por favor espera, estamos recolectando datos para generar tu presentación PDF.';
          showModal(msg);
          try {
            for (let i = 0; i < idChunks.length; i += 1) {
              if (multi && modalText) modalText.textContent = 'Generando parte ' + (i + 1) + ' de ' + idChunks.length + '...';
              const part = 'parte-' + (i + 1) + '-de-' + idChunks.length;
              const url = baseUrl
                + '&pages=' + encodeURIComponent(pages.join(','))
                + '&ids=' + encodeURIComponent(idChunks[i].join(','))
                + '&part=' + encodeURIComponent(part);
              const res = await fetch(url, { credentials: 'same-origin' });
              if (!res.ok) {
                const text = await res.text();
                let errMsg = 'No se pudo generar el PDF.';
                try {
                  const json = JSON.parse(text);
                  errMsg = json.error || json.details || errMsg;
                } catch {
                  if (/502|bad gateway|service is currently unavailable/i.test(text)) {
                    errMsg = 'El servidor tardó demasiado al generar el PDF. Intenta con menos unidades o vuelve a intentar en unos minutos.';
                  }
                }
                throw new Error(errMsg);
              }
              const blob = await res.blob();
              const fallbackName = lang === 'en' ? ('Solar-Midtown-EN-' + part + '.pdf') : ('Solar-Midtown-ES-' + part + '.pdf');
              const fileName = parseFileName(res.headers.get('content-disposition'), fallbackName);
              const objUrl = URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = objUrl;
              a.download = fileName;
              document.body.appendChild(a);
              a.click();
              a.remove();
              URL.revokeObjectURL(objUrl);
            }
            if (multi) alert('Descarga completada en ' + idChunks.length + ' archivos PDF.');
          } catch (err) {
            alert((err && err.message) ? err.message : 'No se pudo descargar el PDF.');
          } finally {
            hideModal();
          }
        }
        function selectedPages() {
          return checks
            .filter((c) => c.checked)
            .map((c) => Number(c.value))
            .filter((n) => Number.isFinite(n) && n > 0)
            .sort((a, b) => a - b);
        }
        function updateDownloadLink() {
          const pages = selectedPages();
          if (hint) hint.textContent = 'Seleccionadas: ' + pages.length + ' de ' + checks.length;
          if (!pages.length) {
            downloadBtnEs.setAttribute('href', '#');
            downloadBtnEn.setAttribute('href', '#');
            downloadBtnEs.style.opacity = '0.5';
            downloadBtnEn.style.opacity = '0.5';
            return;
          }
          downloadBtnEs.style.opacity = '1';
          downloadBtnEn.style.opacity = '1';
          downloadBtnEs.setAttribute('href', baseUrlEs + '&pages=' + encodeURIComponent(pages.join(',')));
          downloadBtnEn.setAttribute('href', baseUrlEn + '&pages=' + encodeURIComponent(pages.join(',')));
        }
        checks.forEach((check) => check.addEventListener('change', updateDownloadLink));
        if (selectAllBtn) selectAllBtn.addEventListener('click', function () {
          checks.forEach((check) => { check.checked = true; });
          updateDownloadLink();
        });
        if (selectNoneBtn) selectNoneBtn.addEventListener('click', function () {
          checks.forEach((check) => { check.checked = false; });
          updateDownloadLink();
        });
        function onDownloadClick(evt) {
          evt.preventDefault();
          const lang = evt.currentTarget === downloadBtnEn ? 'en' : 'es';
          downloadCurrent(lang);
        }
        downloadBtnEs.addEventListener('click', onDownloadClick);
        downloadBtnEn.addEventListener('click', onDownloadClick);
        updateDownloadLink();
      })();
    </script>`;
  return res.send(buildSolarMidtownPagesDocument({
    title: 'Vista previa · Solar Midtown',
    contentTop,
    pagesHtml,
    layout,
    renderProfile
  }));
});

app.get('/api/presentaciones/solar-midtown/download.pdf', async (req, res) => {
  try {
    const lang = normalizeSolarMidtownLang(req.query && req.query.lang);
    const renderProfile = getSolarMidtownPdfRenderProfile(req.query && req.query.quality);
    const brochurePath = getSolarMidtownBrochurePath(lang);
    const data = readSolarMidtownRowsFromCbs();
    const selectedIds = parseSolarMidtownSelectedIds(req.query && req.query.ids, data.rows);
    if (!selectedIds.length) {
      return res.status(400).json({ error: 'Selecciona al menos una unidad para descargar.' });
    }
    const baseBytes = await readSolarMidtownBrochureBytes(brochurePath);
    if (!baseBytes) {
      return res.status(404).json({
        error: 'No se encontró brochure local de Solar Midtown.',
        expectedPath: brochurePath
      });
    }
    const layout = readSolarMidtownLayout();
    const selectedRows = data.rows.filter((row) => selectedIds.includes(row.id));
    if (!selectedRows.length) {
      return res.status(400).json({ error: 'No se encontraron filas válidas para las unidades seleccionadas.' });
    }

    const baseDoc = await PDFDocument.load(baseBytes);
    const selectedBasePageIndices = parseSolarMidtownBrochurePages(req.query && req.query.pages, baseDoc.getPageCount());
    if (!selectedBasePageIndices.length) {
      return res.status(400).json({ error: 'Selecciona al menos 1 hoja del brochure para descargar.' });
    }
    const baseFirstPage = baseDoc.getPage(0);
    const pageWidthPt = baseFirstPage.getWidth();
    const pageHeightPt = baseFirstPage.getHeight();

    const merged = await PDFDocument.create();
    const basePages = await merged.copyPages(baseDoc, selectedBasePageIndices);
    basePages.forEach((page) => merged.addPage(page));
    const pageWidthIn = (pageWidthPt / 72).toFixed(4);
    const pageHeightIn = (pageHeightPt / 72).toFixed(4);
    const rowChunks = splitIntoChunks(selectedRows, SOLAR_MIDTOWN_APPENDIX_CHUNK_SIZE);
    for (const rowChunk of rowChunks) {
      const planDataUrlById = {};
      await mapWithConcurrency(rowChunk, SOLAR_MIDTOWN_PLAN_FETCH_CONCURRENCY, async (row) => {
        const rowId = String(row && row.id || '');
        const link = String(row && row.planLink || '').trim();
        if (!rowId || !link) return;
        const dataUrl = await fetchImageAsDataUrl(link);
        if (dataUrl) planDataUrlById[rowId] = dataUrl;
      });
      const appendixPagesHtml = buildSolarMidtownPagesHtml(rowChunk, layout, {
        absoluteAssets: true,
        planDataUrlById,
        lang
      });
      const appendixHtml = buildSolarMidtownPagesDocument({
        title: 'Solar Midtown · Páginas agregadas',
        contentTop: '',
        pagesHtml: appendixPagesHtml,
        pdfMode: true,
        pageWidthPt,
        pageHeightPt,
        layout,
        renderProfile
      });
      const appendixPdfBuffer = await generatePdfWithSharedBrowser(appendixHtml, {
        width: `${pageWidthIn}in`,
        height: `${pageHeightIn}in`,
        margin: { top: '0mm', right: '0mm', bottom: '0mm', left: '0mm' }
      });
      const appendixDoc = await PDFDocument.load(appendixPdfBuffer);
      const appendixPages = await merged.copyPages(appendixDoc, appendixDoc.getPageIndices());
      appendixPages.forEach((page) => merged.addPage(page));
    }
    const finalPdf = await merged.save();

    const fileName = lang === 'en' ? 'SOLAR MT ENG.pdf' : 'SOLAR MT ESP.pdf';
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    return res.send(Buffer.from(finalPdf));
  } catch (err) {
    log(`Error en /api/presentaciones/solar-midtown/download.pdf: ${err && err.stack ? err.stack : err}`);
    return res.status(500).json({
      error: 'No se pudo generar la presentación PDF',
      details: err && err.message ? err.message : 'error desconocido'
    });
  }
});

app.get('/viceroy', (req, res) => {
  const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
  const showMartinCards = currentEmail === GERENTE_EMAIL;
  const tablaPagosViceroyCard = `
        <a class="card" href="/viceroy/tabla-pagos">
          <span class="tag">Módulo</span>
          <h2 class="name">Tabla de Pago Viceroy</h2>
          <p class="desc">Versión comercial Viceroy basada en acuerdos de RELATED.</p>
        </a>
        ${showMartinCards ? `
        <a class="card" href="/viceroy/tabla-pagos/editor">
          <span class="tag">Editor</span>
          <h2 class="name">Editor PDF Viceroy</h2>
          <p class="desc">Ajusta márgenes, escalas, líneas y orden visual de la hoja final.</p>
        </a>` : ''}`;
  const inventoryCard = showMartinCards ? `
        <a class="card" href="/viceroy-piloto">
          <span class="tag">Módulo</span>
          <h2 class="name">Edición Viceroy Inventario</h2>
          <p class="desc">Flujo visual de tipologías e inventario por piso.</p>
        </a>` : '';
  const inicioCard = showMartinCards ? `
        <a class="card" href="/viceroy/inicio">
          <span class="tag">Módulo</span>
          <h2 class="name">Viceroy Inicio</h2>
          <p class="desc">Acceso al módulo de presentación comercial.</p>
        </a>` : '';
  res.send(`<!doctype html>
  <html lang="es"><head><meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Viceroy</title>
  <style>
    :root{--bg:#f4f1e8;--card:#ffffff;--ink:#1a1a1a;--muted:#5f5f5f;--accent:#ffe816;--line:#d8d1c1;}
    *{box-sizing:border-box}
    body{font-family: Arial, sans-serif; margin:0; background:var(--bg); color:var(--ink);}
    .wrap{max-width:900px; margin:0 auto; padding:28px 18px 48px;}
    .top{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:18px}
    .back{display:inline-block;padding:8px 10px;border:1px solid #bdb8a9;border-radius:10px;background:#fff;color:#111;text-decoration:none;font-size:13px;font-weight:700;}
    h1{margin:0;font-size:30px}
    .sub{margin:6px 0 0;color:var(--muted)}
    .grid{display:grid;grid-template-columns:repeat(3,minmax(220px,1fr));gap:14px}
    .card{display:block;background:#fff;border:1px solid #dcd7cb;border-radius:14px;padding:18px;text-decoration:none;color:inherit}
    .card:hover{border-color:#b9b39f}
    .tag{display:inline-block;font-size:12px;font-weight:700;background:var(--accent);padding:4px 8px;border-radius:999px;margin-bottom:10px}
    .name{font-size:20px;margin:0 0 8px}
    .desc{margin:0;color:var(--muted)}
    @media (max-width:760px){.grid{grid-template-columns:1fr} h1{font-size:24px}}
  </style></head><body>
    <div class="wrap">
      <div class="top">
        <div>
          <h1>Viceroy</h1>
          <p class="sub">Módulos internos de Viceroy.</p>
        </div>
        <a class="back" href="/">Volver al backend</a>
      </div>
      <div class="grid">
        ${inicioCard}
        <a class="card" href="/whisperlist">
          <span class="tag">Módulo</span>
          <h2 class="name">Viceroy Whisperlist</h2>
          <p class="desc">Módulo original de Whisperlist.</p>
        </a>
        <a class="card" href="/viceroy/reservas">
          <span class="tag">Módulo</span>
          <h2 class="name">Reserva de oficinas</h2>
          <p class="desc">Reserva sala grande o sala chica por horario.</p>
        </a>
        <a class="card" href="/viceroy/registros">
          <span class="tag">Módulo</span>
          <h2 class="name">Viceroy Registros</h2>
          <p class="desc">Duplicado de Whisperlist para operación separada.</p>
        </a>
        ${tablaPagosViceroyCard}
        ${inventoryCard}
      </div>
    </div>
  </body></html>`);
});

app.get('/viceroy/reservas', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-reservas.html'));
});

app.get('/api/viceroy/reservas', (req, res) => {
  const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
  const currentName = String(req.user && req.user.name || '').trim();
  const isGerente = currentEmail === GERENTE_EMAIL;
  const from = normalizeReservationDate(req.query && req.query.from);
  const to = normalizeReservationDate(req.query && req.query.to);
  const data = readViceroyRoomReservations();
  const rows = data.rows.filter((row) => {
    if (from && row.date < from) return false;
    if (to && row.date > to) return false;
    return true;
  });
  return res.json({
    ok: true,
    currentEmail,
    currentName,
    isGerente,
    updatedAt: data.updatedAt,
    rows
  });
});

app.post('/api/viceroy/reservas', (req, res) => {
  const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
  const currentName = String(req.user && req.user.name || '').trim();
  const body = req.body || {};
  const date = normalizeReservationDate(body.date);
  const room = normalizeReservationRoom(body.room);
  const hour = normalizeReservationHour(body.hour);
  const title = normalizeReservationTitle(body.title);

  if (!date || !room || !hour || !title) {
    return res.status(400).json({ error: 'Faltan datos obligatorios de la reserva.' });
  }

  const data = readViceroyRoomReservations();
  const occupied = data.rows.find((row) => row.date === date && row.room === room && row.hour === hour);
  if (occupied) {
    return res.status(409).json({
      error: 'Horario ocupado',
      row: occupied
    });
  }

  const newRow = normalizeRoomReservationRow({
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    date,
    room,
    hour,
    title,
    createdByEmail: currentEmail,
    createdByName: currentName,
    updatedAt: new Date().toISOString()
  });
  data.rows.push(newRow);
  saveViceroyRoomReservations(data.rows);
  return res.status(201).json({ ok: true, row: newRow });
});

app.delete('/api/viceroy/reservas/:id', (req, res) => {
  const rowId = String(req.params.id || '').trim();
  if (!rowId) return res.status(400).json({ error: 'id inválido' });

  const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
  const isGerente = currentEmail === GERENTE_EMAIL;
  const data = readViceroyRoomReservations();
  const index = data.rows.findIndex((row) => String(row.id || '') === rowId);
  if (index < 0) return res.status(404).json({ error: 'Reserva no encontrada' });

  const target = data.rows[index];
  if (!isGerente && String(target.createdByEmail || '').trim().toLowerCase() !== currentEmail) {
    return res.status(403).json({ error: 'Solo puedes borrar tus propias reservas' });
  }

  data.rows.splice(index, 1);
  saveViceroyRoomReservations(data.rows);
  return res.json({ ok: true });
});

app.get('/viceroy/inicio', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-inicio.html'));
});

app.get('/viceroy/inicio/presentacion', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-inicio-presentacion.html'));
});

app.get('/viceroy/inicio/acabados', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-inicio-acabados.html'));
});

app.get('/viceroy/inicio/planos', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-inicio-planos.html'));
});

app.get('/viceroy/inicio/carpeta-legal', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-inicio-carpeta-legal.html'));
});

app.get('/viceroy/inicio/ubicacion', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-inicio-ubicacion.html'));
});

app.get('/viceroy/inicio/apartar-unidad', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-inicio-apartar-unidad.html'));
});

app.get('/viceroy/inicio/presentacion/pdf', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'assets', 'viceroy', 'maresol-ppt-sales-desktop.pdf'));
});

app.get('/viceroy/registros', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-registros.html'));
});

app.get('/viceroy/registros/qr', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'viceroy-registros-qr.html'));
});

app.get('/api/viceroy/registros', async (req, res) => {
  try {
    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const isGerente = currentEmail === GERENTE_EMAIL;
    const data = await readViceroyRegistrosData();
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
    return res.json({
      ok: true,
      currentEmail,
      isGerente,
      updatedAt: data.updatedAt,
      sourceFile: data.sourceFile,
      rows
    });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo leer Viceroy Registros' });
  }
});

app.get('/api/viceroy/registros/qr-codes', requireGerente, async (req, res) => {
  try {
    const data = await readViceroyRegistrosData();
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

app.get('/api/viceroy/registros/export.xlsx', requireGerente, async (req, res) => {
  try {
    const data = await readViceroyRegistrosData();
    const rows = viceroyRegistrosExportRows(data.rows);
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'ViceroyRegistros');
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    const dateTag = new Date().toISOString().slice(0, 10);
    const fileName = `viceroy-registros-backup-${dateTag}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    return res.send(buffer);
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo exportar Viceroy Registros' });
  }
});

app.patch('/api/viceroy/registros/rows/:id', async (req, res) => {
  try {
    const rowId = String(req.params.id || '').trim();
    if (!rowId) return res.status(400).json({ error: 'id de fila inválido' });

    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const isGerente = currentEmail === GERENTE_EMAIL;
    const data = await readViceroyRegistrosData();
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
      pais: normalizeWhisperlistCountry(body.pais !== undefined ? body.pais : target.pais),
      ciudad: normalizeWhisperlistCity(body.ciudad !== undefined ? body.ciudad : target.ciudad),
      clientEmail: normalizeClientEmail(body.clientEmail !== undefined ? body.clientEmail : target.clientEmail),
      clientPhone: normalizeClientPhone(body.clientPhone !== undefined ? body.clientPhone : target.clientPhone),
      updatedAt: new Date().toISOString()
    };
    data.rows[index] = nextRow;
    await saveViceroyRegistrosRows(data.rows, data.sourceFile || path.basename(VICEROY_REGISTROS_EXCEL_PATH));
    return res.json({ ok: true, row: nextRow });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo guardar fila' });
  }
});

app.post('/api/viceroy/registros/rows', async (req, res) => {
  try {
    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const fallbackName = String(req.user && req.user.name || '').trim();
    const data = await readViceroyRegistrosData();
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
      pais: normalizeWhisperlistCountry(body.pais),
      ciudad: normalizeWhisperlistCity(body.ciudad),
      clientEmail: normalizeClientEmail(body.clientEmail),
      clientPhone: normalizeClientPhone(body.clientPhone),
      kpi: normalizeWhisperlistKpi(body.kpi),
      updatedAt: new Date().toISOString()
    };
    data.rows.push(newRow);
    await saveViceroyRegistrosRows(data.rows, data.sourceFile || path.basename(VICEROY_REGISTROS_EXCEL_PATH));
    return res.status(201).json({ ok: true, row: newRow });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo agregar fila' });
  }
});

app.delete('/api/viceroy/registros/rows/:id', async (req, res) => {
  try {
    const rowId = String(req.params.id || '').trim();
    if (!rowId) return res.status(400).json({ error: 'id de fila inválido' });

    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const isGerente = currentEmail === GERENTE_EMAIL;
    const data = await readViceroyRegistrosData();
    const index = data.rows.findIndex((row) => String(row.id || '') === rowId);
    if (index < 0) return res.status(404).json({ error: 'Fila no encontrada' });

    const target = data.rows[index];
    const ownerEmail = String(target.correo || '').trim().toLowerCase();
    if (!isGerente && ownerEmail !== currentEmail) {
      return res.status(403).json({ error: 'Solo puedes eliminar filas asignadas a tu correo' });
    }

    data.rows.splice(index, 1);
    await saveViceroyRegistrosRows(data.rows, data.sourceFile || path.basename(VICEROY_REGISTROS_EXCEL_PATH));
    return res.json({ ok: true });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo eliminar fila' });
  }
});

app.post('/api/viceroy/registros/rows/:id/move-to-whisperlist', async (req, res) => {
  try {
    const rowId = String(req.params.id || '').trim();
    if (!rowId) return res.status(400).json({ error: 'id de fila inválido' });

    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const isGerente = currentEmail === GERENTE_EMAIL;

    const registrosData = await readViceroyRegistrosData();
    const index = registrosData.rows.findIndex((row) => String(row.id || '') === rowId);
    if (index < 0) return res.status(404).json({ error: 'Fila no encontrada' });

    const source = registrosData.rows[index];
    const ownerEmail = String(source.correo || '').trim().toLowerCase();
    if (!isGerente && ownerEmail !== currentEmail) {
      return res.status(403).json({ error: 'Solo puedes mover filas asignadas a tu correo' });
    }

    const body = req.body || {};
    const candidate = {
      canal: body.canal !== undefined ? body.canal : source.canal,
      tipoVenta: body.tipoVenta !== undefined ? body.tipoVenta : source.tipoVenta,
      nombreCliente: body.nombreCliente !== undefined ? body.nombreCliente : source.nombreCliente,
      recamaras: body.recamaras !== undefined ? body.recamaras : source.recamaras,
      pais: body.pais !== undefined ? body.pais : source.pais,
      ciudad: body.ciudad !== undefined ? body.ciudad : source.ciudad
    };
    const required = {
      canal: normalizeWhisperlistCanal(candidate.canal),
      tipoVenta: normalizeWhisperlistTipoVenta(candidate.tipoVenta),
      nombreCliente: normalizeWhisperlistPersonText(candidate.nombreCliente),
      recamaras: normalizeWhisperlistRecamaras(candidate.recamaras),
      pais: normalizeWhisperlistCountry(candidate.pais),
      ciudad: normalizeWhisperlistCity(candidate.ciudad)
    };
    const missing = Object.entries(required).filter(([, value]) => !String(value || '').trim()).map(([key]) => key);
    if (missing.length) {
      return res.status(400).json({
        error: 'Completa todos los campos requeridos antes de mover: canal, tipoVenta, nombreCliente, recamaras, pais, ciudad'
      });
    }

    const whisperData = await readWhisperlistData();

    const movedRow = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      asesor: normalizeWhisperlistAsesor(source.asesor),
      correo: String(source.correo || '').trim().toLowerCase(),
      canal: required.canal,
      tipoVenta: required.tipoVenta,
      nombreCliente: required.nombreCliente,
      recamaras: required.recamaras,
      pais: required.pais,
      ciudad: required.ciudad,
      clientEmail: normalizeClientEmail(source.clientEmail),
      clientPhone: normalizeClientPhone(source.clientPhone),
      kpi: normalizeWhisperlistKpi(source.kpi),
      updatedAt: new Date().toISOString()
    };

    whisperData.rows.push(movedRow);
    registrosData.rows.splice(index, 1);

    await saveWhisperlistRows(whisperData.rows, whisperData.sourceFile || path.basename(WHISPERLIST_EXCEL_PATH));
    await saveViceroyRegistrosRows(registrosData.rows, registrosData.sourceFile || path.basename(VICEROY_REGISTROS_EXCEL_PATH));
    return res.json({ ok: true, row: movedRow });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo mover a Whisperlist' });
  }
});

app.post('/api/viceroy/registros/import-excel', requireGerente, async (req, res) => {
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

  const parsed = parseViceroyRegistrosWorkbook(workbook, sheetName);
  if (!parsed.sheetName) {
    return res.status(400).json({ error: 'No se encontró una hoja válida en el Excel' });
  }

  const safeName = sanitizeExcelFileName(fileName) || 'viceroy-registros.xlsx';
  const current = await readViceroyRegistrosData();
  const shouldReplace = Boolean(replaceExisting);
  let created = 0;
  let skipped = 0;
  const nextRows = shouldReplace ? [] : [...current.rows];
  const seen = new Set(
    nextRows.map((row) => [
      String(row.correo || '').trim().toLowerCase(),
      String(row.asesor || '').trim().toUpperCase(),
      String(row.nombreCliente || '').trim().toUpperCase(),
      String(row.canal || '').trim().toUpperCase()
      ,
      String(row.tipoVenta || '').trim().toUpperCase()
    ].join('|'))
  );
  parsed.rows.forEach((row) => {
    const key = [
      String(row.correo || '').trim().toLowerCase(),
      String(row.asesor || '').trim().toUpperCase(),
      String(row.nombreCliente || '').trim().toUpperCase(),
      String(row.canal || '').trim().toUpperCase()
      ,
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

  await saveViceroyRegistrosRows(nextRows, safeName);
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

app.get('/api/whisperlist/sellers', async (req, res) => {
  try {
    const data = await readWhisperlistData();
    const sellers = getWhisperlistSellers(data.rows);
    return res.json({ ok: true, sellers });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudieron leer vendedores de whisperlist' });
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
      pais: normalizeWhisperlistCountry(body.pais !== undefined ? body.pais : target.pais),
      ciudad: normalizeWhisperlistCity(body.ciudad !== undefined ? body.ciudad : target.ciudad),
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

app.post('/api/whisperlist/rows/:id/kpi', async (req, res) => {
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
      return res.status(403).json({ error: 'Solo puedes editar KPI de filas asignadas a tu correo' });
    }

    const body = req.body || {};
    const nextRow = {
      ...target,
      pais: normalizeWhisperlistCountry(body.pais !== undefined ? body.pais : target.pais),
      ciudad: normalizeWhisperlistCity(body.ciudad !== undefined ? body.ciudad : target.ciudad),
      kpi: normalizeWhisperlistKpi({
        ...(target.kpi && typeof target.kpi === 'object' ? target.kpi : {}),
        ...(body && typeof body === 'object' ? body : {})
      }),
      updatedAt: new Date().toISOString()
    };
    data.rows[index] = nextRow;
    await saveWhisperlistRows(data.rows, data.sourceFile || path.basename(WHISPERLIST_EXCEL_PATH));
    return res.json({ ok: true, row: nextRow });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo guardar KPI' });
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
      pais: normalizeWhisperlistCountry(body.pais),
      ciudad: normalizeWhisperlistCity(body.ciudad),
      clientEmail: normalizeClientEmail(body.clientEmail),
      clientPhone: normalizeClientPhone(body.clientPhone),
      kpi: normalizeWhisperlistKpi(body.kpi),
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

app.post('/api/whisperlist/rows/:id/send-to-registros', async (req, res) => {
  try {
    const rowId = String(req.params.id || '').trim();
    if (!rowId) return res.status(400).json({ error: 'id de fila inválido' });

    const currentEmail = String(req.user && req.user.email || '').trim().toLowerCase();
    const isGerente = currentEmail === GERENTE_EMAIL;
    const whisperData = await readWhisperlistData();
    const index = whisperData.rows.findIndex((row) => String(row.id || '') === rowId);
    if (index < 0) return res.status(404).json({ error: 'Fila no encontrada' });

    const source = whisperData.rows[index];
    const ownerEmail = String(source.correo || '').trim().toLowerCase();
    if (!isGerente && ownerEmail !== currentEmail) {
      return res.status(403).json({ error: 'Solo puedes enviar filas asignadas a tu correo' });
    }

    const registrosData = await readViceroyRegistrosData();
    const newRow = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      asesor: normalizeWhisperlistAsesor(source.asesor),
      correo: String(source.correo || '').trim().toLowerCase(),
      canal: normalizeWhisperlistCanal(source.canal),
      tipoVenta: normalizeWhisperlistTipoVenta(source.tipoVenta),
      nombreCliente: normalizeWhisperlistPersonText(source.nombreCliente),
      recamaras: normalizeWhisperlistRecamaras(source.recamaras),
      pais: normalizeWhisperlistCountry(source.pais),
      ciudad: normalizeWhisperlistCity(source.ciudad),
      clientEmail: normalizeClientEmail(source.clientEmail),
      clientPhone: normalizeClientPhone(source.clientPhone),
      updatedAt: new Date().toISOString()
    };

    registrosData.rows.push(newRow);
    whisperData.rows.splice(index, 1);

    await saveViceroyRegistrosRows(registrosData.rows, registrosData.sourceFile || path.basename(VICEROY_REGISTROS_EXCEL_PATH));
    await saveWhisperlistRows(whisperData.rows, whisperData.sourceFile || path.basename(WHISPERLIST_EXCEL_PATH));
    return res.json({ ok: true, row: newRow });
  } catch (err) {
    return res.status(500).json({ error: 'No se pudo enviar a Registros' });
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
      pais: normalizeWhisperlistCountry(req.body && req.body.pais),
      ciudad: normalizeWhisperlistCity(req.body && req.body.ciudad),
      clientEmail,
      clientPhone,
      kpi: normalizeWhisperlistKpi({}),
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
  const config = readViceroyPilotoConfig();
  const requestedName = typeof req.query.name === 'string' ? req.query.name.trim() : '';
  const selectedName = requestedName || String(config.selectedFloorJsonName || '').trim();
  let floorsData = selectedName
    ? readNamedFloorsByDevelopment(devSlug, selectedName)
    : readMergedFloorsByDevelopment(devSlug);
  if ((!Array.isArray(floorsData.floors) || !floorsData.floors.length) && selectedName) {
    floorsData = readMergedFloorsByDevelopment(devSlug);
  }
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
  let saved = updateViceroyPilotoPresentationLayout(layoutPayload);
  if (Object.prototype.hasOwnProperty.call(body, 'selectedFloorJsonName')) {
    saved = updateViceroyPilotoSelectedFloorJson(body.selectedFloorJsonName);
  }
  return res.json({
    ok: true,
    presentationLayout: saved.presentationLayout,
    selectedFloorJsonName: saved.selectedFloorJsonName || ''
  });
});

app.get('/api/viceroy-piloto/floors', (req, res) => {
  const devSlug = 'viceroy-piloto';
  const config = readViceroyPilotoConfig();
  const requestedName = typeof req.query.name === 'string' ? req.query.name.trim() : '';
  const selectedName = requestedName || String(config.selectedFloorJsonName || '').trim();
  let data = selectedName
    ? readNamedFloorsByDevelopment(devSlug, selectedName)
    : readMergedFloorsByDevelopment(devSlug);
  if ((!Array.isArray(data.floors) || !data.floors.length) && selectedName) {
    data = readMergedFloorsByDevelopment(devSlug);
  }
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
  const allowDownloadsFallback = dev.slug === DEFAULT_DEVELOPMENT_SLUG;
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
    if (!filePath && allowDownloadsFallback) {
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

  if (!filePath && safeName && allowDownloadsFallback) {
    try {
      const files = fs.readdirSync(downloadsDir);
      const preferred = files.find((name) => name.toLowerCase() === safeName.toLowerCase());
      if (preferred) filePath = path.join(downloadsDir, preferred);
    } catch {}
  }

  if (!filePath && !safeName && allowDownloadsFallback) {
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
  const allowDownloadsFallback = dev.slug === DEFAULT_DEVELOPMENT_SLUG;
  const floorDirs = getDevelopmentFloorSearchDirs(dev.slug);
  try {
    for (const dir of floorDirs) {
      const files = listFloorJsonFiles(dir);
      if (files.length) {
        return res.json({ files, sourceDir: dir, dev: dev.slug });
      }
    }

    if (!allowDownloadsFallback) {
      return res.json({ files: [], sourceDir: floorDirs.join(', '), dev: dev.slug });
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
    const safePayload = {
      ...payload,
      developmentSlug: dev.slug
    };
    fs.writeFileSync(filePath, JSON.stringify(safePayload, null, 2), 'utf-8');
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

app.get('/form-extranjera-moral', (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, 'form-extranjera-moral.html'));
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

app.get('/api/plds/formats-extranjera-moral', (req, res) => {
  const items = Object.entries(foreignMoralFormats).map(([id, f]) => ({
    id,
    name: f.name,
    template: f.file,
    htmlRoute: `/format-extranjera-moral/${id}`,
    pdfGetRoute: `/format-extranjera-moral/${id}/pdf`,
    pdfPostRoute: `/format-extranjera-moral/${id}/pdf`
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
  if (id === '10' || id === '19' || id === '21' || id === '26' || id === '35') {
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
  if (id === '10' || id === '19' || id === '21' || id === '26' || id === '35') {
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
  if (id === '10' || id === '19' || id === '21' || id === '26' || id === '35') {
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

app.get('/format-extranjera-moral/:id', (req, res) => {
  const id = req.params.id;
  const format = foreignMoralFormats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const data = fs.existsSync(DATA_PATH) ? JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) : {};
  const html = renderTemplate(format.file, data, {
    baseUrl,
    showControls: !req.query.print,
    downloadUrl: `/format-extranjera-moral/${id}/pdf`,
    bodyClass: 'interactive'
  });
  res.send(html);
});

app.get('/format-extranjera-moral/:id/pdf', async (req, res) => {
  const id = req.params.id;
  const format = foreignMoralFormats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const data = fs.existsSync(DATA_PATH) ? JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8')) : {};
  const html = renderTemplate(format.file, data, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format-extranjera-moral/${id}/pdf`,
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
    log(`Error en GET /format-extranjera-moral/${id}/pdf: ${err && err.stack ? err.stack : err}`);
    res.status(500).send('No se pudo generar el PDF');
  }
});

app.post('/format-extranjera-moral/:id/pdf', async (req, res) => {
  const id = req.params.id;
  const format = foreignMoralFormats[id];
  if (!format) return res.status(404).send('Formato no encontrado');

  const data = { ...(req.body || {}) };
  if (id === '10' || id === '19' || id === '21' || id === '26' || id === '35') {
    Object.keys(data).forEach((k) => {
      const v = data[k];
      if (typeof v === 'string' && v.trim() === '') data[k] = 'N/A';
    });
  }

  const saved = persistSubmission(`EM-${id}`, format.name, data);
  log(`Submission guardado: ${saved.id} formato=EM-${id}`);

  const baseUrl = `${req.protocol}://${req.get('host')}`;
  const html = renderTemplate(format.file, data, {
    baseUrl,
    showControls: false,
    downloadUrl: `/format-extranjera-moral/${id}/pdf`,
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
    log(`Error en POST /format-extranjera-moral/${id}/pdf: ${err && err.stack ? err.stack : err}`);
    res.status(500).send('No se pudo generar el PDF');
  }
});

async function startServer() {
  ensureWhisperlistStorageReady()
    .then(() => {
      log('Whisperlist storage listo');
    })
    .catch((err) => {
      log(`Fallo inicializando Whisperlist storage: ${err && err.stack ? err.stack : err}`);
    });
  const server = app.listen(PORT, HOST, () => {
    log(`Servidor listo en http://${HOST}:${PORT}`);
  });
  // Helps prevent intermittent gateway resets on long-running PDF generation requests.
  server.keepAliveTimeout = 120000;
  server.headersTimeout = 121000;
  server.requestTimeout = 300000;
  server.on('error', (err) => {
    log(`Error al iniciar servidor: ${err && err.stack ? err.stack : err}`);
  });
}

startServer();
