// End-to-end regression for the ported PPTX conversion chain.
//
//   node scripts/verify-convert.mjs
//
// Bundles the host-side glue (src/convert.ts -> parser's @fefeding/ppt-parser/convert) the same way esbuild.js does for
// dist/extension.js, then converts every examples/*.pptx to docx / pdf / md / html and checks
// the results are real documents (magic bytes / markup), not empty or partially failed output.
// No VS Code instance is needed: neither module imports `vscode`.

import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import { readdirSync, readFileSync } from 'fs';
import path from 'path';
import os from 'os';

const require = createRequire(import.meta.url);
const esbuild = require('esbuild');
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outFile = path.join(os.tmpdir(), `vscode-pptx-convert-${process.pid}.cjs`);

// The extension now lives inside the pptx-parser repo. Resolve @fefeding/ppt-parser
// straight to the parent repo's TypeScript source (as esbuild.js does) instead of npm.
const parserRoot = path.resolve(root, '..');

await esbuild.build({
  entryPoints: [path.join(root, 'src', 'convert.ts')],
  outfile: outFile,
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'node16',
  logLevel: 'silent',
  legalComments: 'none',
  alias: {
    '@fefeding/ppt-parser/convert': path.join(parserRoot, 'src', 'convert', 'index.ts'),
    '@fefeding/ppt-parser': path.join(parserRoot, 'src', 'index.ts'),
    '@fefeding/ppt-parser/chart-renderer': path.join(parserRoot, 'examples', 'chart-lib', 'chart-renderer.js')
  },
  loader: { '.css': 'text' },
  nodePaths: [path.join(root, 'node_modules'), path.join(parserRoot, 'node_modules')]
});

const { CONVERT_TARGETS, convertPptxBytes, isPptxPath } = require(outFile);

let checks = 0;
function assert(ok, message) {
  checks++;
  if (!ok) throw new Error(message);
}

assert(CONVERT_TARGETS.length === 4, 'expected 4 conversion targets');
assert(CONVERT_TARGETS.map((t) => t.id).join(',') === 'docx,pdf,md,html', 'unexpected target order');
assert(isPptxPath('Deck.PPTX') && !isPptxPath('Deck.pdf'), 'isPptxPath regression');

const examples = path.join(root, 'examples');
const files = readdirSync(examples).filter((f) => /\.pptx$/i.test(f));
assert(files.length > 0, 'no example .pptx found in examples/');

for (const name of files) {
  const bytes = new Uint8Array(readFileSync(path.join(examples, name)));
  const title = name.replace(/\.[^.]+$/, '');
  for (const meta of CONVERT_TARGETS) {
    const out = await convertPptxBytes(bytes, meta.id, title);
    assert(out != null, `${name} -> ${meta.id}: no result`);
    const size = typeof out === 'string' ? Buffer.byteLength(out) : out.length;
    assert(size > 200, `${name} -> ${meta.id}: result too small (${size} bytes)`);
    if (meta.id === 'docx') {
      const head = Buffer.from(out.slice(0, 4)).toString('hex');
      assert(head === '504b0304', `${name} -> docx: not a ZIP (${head})`);
    } else if (meta.id === 'pdf') {
      assert(Buffer.from(out.slice(0, 5)).toString() === '%PDF-', `${name} -> pdf: missing %PDF- header`);
    } else if (meta.id === 'md') {
      assert(typeof out === 'string' && out.trim().length > 0, `${name} -> md: empty markdown`);
    } else {
      assert(
        typeof out === 'string' && out.startsWith('<!DOCTYPE html>') && out.includes('jt-slide'),
        `${name} -> html: missing slide markup`
      );
    }
    console.log(`  ${name} -> ${meta.id.padEnd(4)} ${String(size).padStart(9)} bytes  OK`);
  }
}

// ---- Stage 2: full command flow (QuickPick -> convert -> save dialog -> write) ----
// Bundles the real extension entry with `vscode` external, then drives the registered
// `pptx-parser.convert` command through a stubbed VS Code API.
const extOutFile = path.join(os.tmpdir(), `vscode-pptx-extension-verify-${process.pid}.cjs`);
await esbuild.build({
  entryPoints: [path.join(root, 'src', 'extension.ts')],
  outfile: extOutFile,
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: 'node16',
  external: ['vscode'],
  logLevel: 'silent',
  legalComments: 'none'
});

const Module = require('module');
const originalLoad = Module._load;
const handlers = {};
const written = [];
let pickIndex = 0;
const fakeUri = (p) => ({ path: p, fsPath: p, toString: () => `file://${p}`, with: (o) => fakeUri(o.path || p) });
const sourceUri = fakeUri('/tmp/deck.pptx');
const vscodeStub = {
  window: {
    registerCustomEditorProvider: () => ({ dispose() {} }),
    activeTextEditor: undefined,
    showQuickPick: async (items) => items[pickIndex],
    showSaveDialog: async () => fakeUri(`/tmp/deck.${CONVERT_TARGETS[pickIndex].ext}`),
    showInformationMessage: () => {},
    showWarningMessage: (m) => { throw new Error(`unexpected warning: ${m}`); },
    showErrorMessage: (m) => { throw new Error(`unexpected error: ${m}`); },
    withProgress: async (_options, fn) => fn()
  },
  commands: {
    registerCommand: (id, fn) => { handlers[id] = fn; return { dispose() {} }; },
    executeCommand: async () => {}
  },
  Uri: { joinPath: () => fakeUri('/tmp'), file: (p) => fakeUri(p) },
  EventEmitter: class { constructor() { this.event = () => ({ dispose() {} }); } fire() {} dispose() {} },
  ProgressLocation: { Notification: 15 },
  env: { language: 'en' },
  workspace: {
    fs: {
      readFile: async () => new Uint8Array(readFileSync(path.join(examples, files[0]))),
      writeFile: async (uri, bytes) => written.push([uri.fsPath, Buffer.from(bytes)])
    }
  },
  l10n: { t: (key) => key }
};
Module._load = function (request, ...rest) {
  if (request === 'vscode') return vscodeStub;
  return originalLoad.call(this, request, ...rest);
};
try {
  require(extOutFile).activate({ subscriptions: [], extensionUri: {}, asAbsolutePath: (p) => p });
  assert(typeof handlers['pptx-parser.convert'] === 'function', 'convert command not registered');
  for (let i = 0; i < CONVERT_TARGETS.length; i++) {
    pickIndex = i;
    written.length = 0;
    const meta = CONVERT_TARGETS[i];
    await handlers['pptx-parser.convert'](sourceUri);
    assert(written.length === 1, `command -> ${meta.id}: nothing was written`);
    const [dest, bytes] = written[0];
    assert(dest.endsWith('.' + meta.ext), `command -> ${meta.id}: wrong output name ${dest}`);
    if (meta.id === 'docx') assert(bytes.slice(0, 4).toString('hex') === '504b0304', 'command -> docx: not a ZIP');
    if (meta.id === 'pdf') assert(bytes.slice(0, 5).toString() === '%PDF-', 'command -> pdf: missing header');
    if (meta.id === 'md') assert(bytes.toString('utf8').trim().length > 0, 'command -> md: empty');
    if (meta.id === 'html') assert(bytes.toString('utf8').includes('jt-slide'), 'command -> html: missing slides');
    console.log(`  command -> ${meta.id.padEnd(4)} ${String(bytes.length).padStart(9)} bytes  OK`);
  }
} finally {
  Module._load = originalLoad;
}

console.log(`verify-convert: ${checks} assertions passed.`);
