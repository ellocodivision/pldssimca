const fs = require('fs');
const path = require('path');
const os = require('os');

const PROJECT_ROOT = path.resolve(__dirname, '..');
const OUTPUT_DIR = path.resolve(
  process.env.VICEROY_LAYOUT_OUTPUT_DIR ||
  path.join(os.tmpdir(), `viceroy-presentation-layouts-${new Date().toISOString().replace(/[:.]/g, '-')}`)
);
const SOURCE_DIR = process.env.VICEROY_LAYOUT_SOURCE_DIR ? path.resolve(process.env.VICEROY_LAYOUT_SOURCE_DIR) : '';
const BASE_URL = String(process.env.VICEROY_BASE_URL || '').trim() || 'http://127.0.0.1:3000';
const SESSION_COOKIE = String(process.env.VICEROY_SESSION_COOKIE || process.env.VICEROY_COOKIE || '').trim();
const MODULES = String(process.env.VICEROY_LAYOUT_MODULES || 'welcome-client,presentation-multi')
  .split(',')
  .map((value) => String(value || '').trim())
  .filter(Boolean);
const LANGS = String(process.env.VICEROY_LAYOUT_LANGS || 'es,en')
  .split(',')
  .map((value) => String(value || '').trim())
  .filter(Boolean);

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function layoutFileName(lang, moduleKey) {
  if (moduleKey === 'presentation-multi') {
    return lang === 'en' ? 'viceroy-presentacion-multi-layout-en.json' : 'viceroy-presentacion-multi-layout.json';
  }
  return lang === 'en' ? 'viceroy-presentacion-layout-en.json' : 'viceroy-presentacion-layout.json';
}

function sourcePathFor(lang, moduleKey) {
  const fileName = layoutFileName(lang, moduleKey);
  return SOURCE_DIR ? path.join(SOURCE_DIR, fileName) : '';
}

async function readLayoutFromSource(lang, moduleKey) {
  const filePath = sourcePathFor(lang, moduleKey);
  if (!filePath || !fs.existsSync(filePath)) return null;
  const raw = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  return {
    ok: true,
    layout: raw,
    defaults: null,
    pageSizes: [],
    module: moduleKey,
    language: lang,
    source: 'local-file'
  };
}

async function fetchLayout(lang, moduleKey) {
  if (SOURCE_DIR) {
    const fromFile = await readLayoutFromSource(lang, moduleKey);
    if (fromFile) return fromFile;
  }
  const url = new URL('/api/viceroy/presentacion-generador/layout', BASE_URL);
  url.searchParams.set('lang', lang);
  url.searchParams.set('module', moduleKey);
  const headers = {};
  if (SESSION_COOKIE) headers.cookie = SESSION_COOKIE;
  const res = await fetch(url, {
    credentials: 'include',
    headers
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data || !data.ok) {
    throw new Error(data && (data.error || data.details) ? (data.error || data.details) : `No se pudo exportar ${moduleKey} (${lang})`);
  }
  return data;
}

async function main() {
  ensureDir(OUTPUT_DIR);
  const bundle = {
    createdAt: new Date().toISOString(),
    projectRoot: PROJECT_ROOT,
    baseUrl: BASE_URL,
    modules: MODULES,
    langs: LANGS,
    items: []
  };

  for (const moduleKey of MODULES) {
    for (const lang of LANGS) {
      const data = await fetchLayout(lang, moduleKey);
      const fileName = layoutFileName(lang, moduleKey);
      const outPath = path.join(OUTPUT_DIR, fileName);
      fs.writeFileSync(outPath, JSON.stringify(data.layout, null, 2) + '\n', 'utf8');
      bundle.items.push({
        module: moduleKey,
        language: lang,
        fileName,
        filePath: outPath,
        pageSizes: Array.isArray(data.pageSizes) ? data.pageSizes.length : 0
      });
    }
  }

  fs.writeFileSync(path.join(OUTPUT_DIR, 'bundle.json'), JSON.stringify(bundle, null, 2) + '\n', 'utf8');
  process.stdout.write([
    `Layouts exported to: ${OUTPUT_DIR}`,
    ...bundle.items.map((item) => `- ${item.module} / ${item.language}: ${item.fileName}`)
  ].join('\n') + '\n');
}

main().catch((err) => {
  process.stderr.write((err && err.stack) ? `${err.stack}\n` : `${err}\n`);
  process.exitCode = 1;
});
