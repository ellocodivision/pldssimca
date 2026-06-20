const fs = require('fs');
const path = require('path');

const PROJECT_ROOT = path.resolve(__dirname, '..');
const INPUT_PATH = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.resolve(process.env.VICEROY_LAYOUT_IMPORT_PATH || '');
const TARGET_DIR = path.join(PROJECT_ROOT, 'data');

const TARGET_FILES = {
  'welcome-client:es': 'viceroy-presentacion-layout.json',
  'welcome-client:en': 'viceroy-presentacion-layout-en.json',
  'presentation-multi:es': 'viceroy-presentacion-multi-layout.json',
  'presentation-multi:en': 'viceroy-presentacion-multi-layout-en.json'
};

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function pickLayoutItem(bundle, moduleKey, language) {
  const items = Array.isArray(bundle && bundle.items) ? bundle.items : [];
  return items.find((item) => String(item.module || '').trim() === moduleKey && String(item.language || '').trim() === language) || null;
}

function loadJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function resolveInputBundle() {
  if (!INPUT_PATH) {
    throw new Error('Pasa una ruta a bundle.json o a una carpeta exportada');
  }
  const stat = fs.statSync(INPUT_PATH);
  if (stat.isDirectory()) {
    return {
      type: 'directory',
      path: INPUT_PATH,
      bundle: fs.existsSync(path.join(INPUT_PATH, 'bundle.json'))
        ? loadJson(path.join(INPUT_PATH, 'bundle.json'))
        : null
    };
  }
  return {
    type: 'file',
    path: INPUT_PATH,
    bundle: loadJson(INPUT_PATH)
  };
}

function resolveLayoutJson(input, moduleKey, language) {
  if (input.type === 'directory') {
    const fileName = TARGET_FILES[`${moduleKey}:${language}`];
    const filePath = path.join(input.path, fileName);
    if (!fs.existsSync(filePath)) {
      throw new Error(`No se encontró ${fileName} en ${input.path}`);
    }
    return loadJson(filePath);
  }

  const bundle = input.bundle;
  const item = pickLayoutItem(bundle, moduleKey, language);
  if (item && item.fileName) {
    const bundleDir = path.dirname(input.path);
    const filePath = path.join(bundleDir, item.fileName);
    if (fs.existsSync(filePath)) {
      return loadJson(filePath);
    }
  }
  const directKey = TARGET_FILES[`${moduleKey}:${language}`];
  if (bundle && bundle[directKey]) return bundle[directKey];
  if (bundle && bundle.layout && typeof bundle.layout === 'object') return bundle.layout;
  throw new Error(`No se pudo resolver ${moduleKey} (${language}) desde el bundle`);
}

function main() {
  const input = resolveInputBundle();
  ensureDir(TARGET_DIR);

  const mappings = [
    ['welcome-client', 'es'],
    ['welcome-client', 'en'],
    ['presentation-multi', 'es'],
    ['presentation-multi', 'en']
  ];

  const written = [];
  for (const [moduleKey, language] of mappings) {
    const layout = resolveLayoutJson(input, moduleKey, language);
    const outName = TARGET_FILES[`${moduleKey}:${language}`];
    const outPath = path.join(TARGET_DIR, outName);
    fs.writeFileSync(outPath, JSON.stringify(layout, null, 2) + '\n', 'utf8');
    written.push(outName);
  }

  process.stdout.write([
    `Layouts importados en: ${TARGET_DIR}`,
    ...written.map((name) => `- ${name}`)
  ].join('\n') + '\n');
}

main();
