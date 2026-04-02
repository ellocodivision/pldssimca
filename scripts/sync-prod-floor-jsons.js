const fs = require('fs');
const path = require('path');

const BASE_URL = String(process.env.SYNC_BASE_URL || 'https://backend-simca.onrender.com').replace(/\/+$/, '');
const SESSION_COOKIE = String(process.env.SYNC_COOKIE || '').trim();
const TARGET_ROOT = path.resolve(process.env.SYNC_TARGET_DIR || path.join(process.cwd(), 'data', 'developments'));
const RUNTIME_ROOT = String(process.env.SYNC_RUNTIME_DIR || '').trim();
const ONLY_DEV = String(process.env.SYNC_DEV || '').trim().toLowerCase();

const DEVELOPMENTS = [
  'ceiba',
  'costa-caribe',
  'cruz-con-mar',
  'dream-c',
  'gran-tulum',
  'ipana',
  'maresol',
  'marila',
  'natal',
  'saint-marine',
  'serenada',
  'singular-joy',
  'solar',
  'solar-mt',
  'viceroy-piloto'
];

if (!SESSION_COOKIE) {
  console.error('Falta SYNC_COOKIE con la cookie de sesión autenticada.');
  process.exit(1);
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function safeJsonName(name) {
  const base = path.basename(String(name || '').trim());
  if (!base || !base.toLowerCase().endsWith('.json')) return '';
  return base;
}

async function fetchJson(url) {
  const res = await fetch(url, {
    headers: {
      Cookie: SESSION_COOKIE,
      Accept: 'application/json'
    }
  });
  const text = await res.text();
  let data = null;
  try {
    data = JSON.parse(text);
  } catch {
    data = null;
  }
  if (!res.ok) {
    const detail = data && data.error ? data.error : text.slice(0, 200);
    throw new Error(`${res.status} ${res.statusText} - ${detail}`);
  }
  return data;
}

async function fetchBuffer(url) {
  const res = await fetch(url, {
    headers: {
      Cookie: SESSION_COOKIE,
      Accept: 'application/json,*/*'
    }
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`${res.status} ${res.statusText} - ${text.slice(0, 200)}`);
  }
  const arr = await res.arrayBuffer();
  return Buffer.from(arr);
}

function writeTargets(dev, fileName, buffer) {
  const relative = path.join(dev, 'plano-ventas-floors', fileName);
  const targets = [path.join(TARGET_ROOT, relative)];
  if (RUNTIME_ROOT) {
    targets.push(path.join(path.resolve(RUNTIME_ROOT), 'developments', relative));
  }
  targets.forEach((targetPath) => {
    ensureDir(path.dirname(targetPath));
    fs.writeFileSync(targetPath, buffer);
  });
}

async function syncDevelopment(dev) {
  const listUrl = `${BASE_URL}/api/plano-ventas/json-files?dev=${encodeURIComponent(dev)}`;
  const data = await fetchJson(listUrl);
  const files = Array.isArray(data && data.files) ? data.files.map(safeJsonName).filter(Boolean) : [];
  if (!files.length) {
    console.log(`[${dev}] sin JSONs listados`);
    return { dev, downloaded: 0 };
  }

  let downloaded = 0;
  for (const fileName of files) {
    const fileUrl = `${BASE_URL}/api/plano-ventas/default-json?dev=${encodeURIComponent(dev)}&name=${encodeURIComponent(fileName)}`;
    const buffer = await fetchBuffer(fileUrl);
    writeTargets(dev, fileName, buffer);
    downloaded += 1;
    console.log(`[${dev}] ${fileName}`);
  }
  return { dev, downloaded };
}

async function main() {
  const devs = ONLY_DEV ? DEVELOPMENTS.filter((dev) => dev === ONLY_DEV) : DEVELOPMENTS.slice();
  if (!devs.length) {
    throw new Error(`Desarrollo inválido en SYNC_DEV: ${ONLY_DEV}`);
  }
  ensureDir(TARGET_ROOT);
  const summary = [];
  for (const dev of devs) {
    try {
      summary.push(await syncDevelopment(dev));
    } catch (err) {
      console.error(`[${dev}] ERROR: ${err && err.message ? err.message : err}`);
    }
  }
  const total = summary.reduce((acc, item) => acc + Number(item && item.downloaded || 0), 0);
  console.log(`Descargados ${total} JSON(s) en ${TARGET_ROOT}`);
  if (RUNTIME_ROOT) {
    console.log(`Runtime sincronizado en ${path.resolve(RUNTIME_ROOT)}`);
  }
}

main().catch((err) => {
  console.error(err && err.message ? err.message : err);
  process.exit(1);
});
