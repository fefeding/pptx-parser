/**
 * Host-side l10n consistency check.
 *
 * The webview locales (src/webview/locales/*.ts) are compile-time checked because zh-cn is typed
 * as Record<keyof typeof en, string>. The VS Code host bundles are plain JSON and had no such
 * guard — a missing `loc('extension.foo')` key silently falls back to the English string, so the
 * previous review only caught one by manual reading. This script closes that gap:
 *   1. every key used through loc(...) / labelKey in src/ exists in bundle.l10n.json;
 *   2. every translated bundle covers the full key set of the English bundle (and no extras);
 *   3. package.nls.json keys match every l10n/<lang>.json command-title bundle;
 *   4. {0}/{1}/... placeholders are identical between English and each translation.
 */
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const src = join(root, 'src');
const l10nDir = join(root, 'l10n');
const errors = [];
const fail = (msg) => errors.push(msg);

const readJson = (p) => JSON.parse(readFileSync(p, 'utf8'));

/** Recursively collect .ts files under src/. */
function tsFiles(dir) {
  const out = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...tsFiles(p));
    else if (entry.name.endsWith('.ts')) out.push(p);
  }
  return out;
}

/** Extract the set of l10n keys referenced from source code. */
function usedKeys() {
  const keys = new Set();
  for (const file of tsFiles(src)) {
    const text = readFileSync(file, 'utf8');
    // loc('extension.foo', ...) — host runtime strings
    for (const m of text.matchAll(/\bloc\(\s*['"]([^'"]+)['"]/g)) keys.add(m[1]);
    // labelKey: 'extension.target.docx' — QuickPick labels from convert.ts
    for (const m of text.matchAll(/\blabelKey\s*:\s*['"]([^'"]+)['"]/g)) keys.add(m[1]);
  }
  return keys;
}

/** Collect {n} placeholders from a string, ignoring order. */
function placeholders(s) {
  return [...String(s).matchAll(/\{(\d+)\}/g)].map(m => m[1]).sort().join(',');
}

const enBundlePath = join(l10nDir, 'bundle.l10n.json');
if (!existsSync(enBundlePath)) {
  fail('l10n/bundle.l10n.json (English fallback) is missing');
} else {
  const en = readJson(enBundlePath);
  const enKeys = Object.keys(en);

  // 1. source references must exist in the English bundle
  for (const key of usedKeys()) {
    if (!(key in en)) fail(`loc()/labelKey references "${key}" but bundle.l10n.json has no such key`);
  }

  // 2. translated bundles must cover the full key set (and add nothing)
  for (const file of readdirSync(l10nDir)) {
    if (!/^[a-z-]+\.bundle\.l10n\.json$/.test(file) || file === 'bundle.l10n.json') continue;
    const dict = readJson(join(l10nDir, file));
    for (const key of enKeys) {
      if (!(key in dict)) fail(`${file}: missing key "${key}"`);
      else if (placeholders(dict[key]) !== placeholders(en[key])) {
        fail(`${file}: key "${key}" placeholders ${placeholders(dict[key]) || '(none)'} != English ${placeholders(en[key]) || '(none)'}`);
      }
    }
    for (const key of Object.keys(dict)) {
      if (!(key in en)) fail(`${file}: extra key "${key}" not present in bundle.l10n.json`);
    }
  }
}

// 3. package.nls.json ↔ l10n/<lang>.json (command titles / editor labels)
const nls = existsSync(join(root, 'package.nls.json')) ? readJson(join(root, 'package.nls.json')) : null;
if (nls) {
  const nlsKeys = Object.keys(nls);
  for (const file of readdirSync(l10nDir)) {
    if (!/^[a-z-]+\.json$/.test(file) || file.includes('.bundle.')) continue;
    const dict = readJson(join(l10nDir, file));
    for (const key of nlsKeys) {
      if (!(key in dict)) fail(`l10n/${file}: missing command-title key "${key}"`);
      else if (placeholders(dict[key]) !== placeholders(nls[key])) {
        fail(`l10n/${file}: key "${key}" placeholders differ from package.nls.json`);
      }
    }
    for (const key of Object.keys(dict)) {
      if (!(key in nls)) fail(`l10n/${file}: extra command-title key "${key}" not in package.nls.json`);
    }
  }
}

if (errors.length) {
  console.error(`check-l10n: ${errors.length} problem(s):`);
  for (const e of errors) console.error(`  - ${e}`);
  process.exit(1);
}
console.log('check-l10n: all host bundles consistent (keys + placeholders).');
