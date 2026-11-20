// 构建脚本：打包扩展宿主（Node）与 Webview（浏览器）两个产物。
//
// 本扩展已从独立仓库合并进 pptx-parser 仓库，位于 <repo>/vscode 子目录。
// 不再从 npm 安装 @fefeding/ppt-parser，而是直接把上层 ../src 的 TypeScript 源码
// 编译并内联进扩展包（extension.js / webview.js），随 VSIX 一起发布。
//
// 用法：
//   node esbuild.js              单次构建（npm run compile）
//   node esbuild.js --watch      持续构建（npm run dev / npm run watch）
//   node esbuild.js --production 生产构建（npm run vscode:prepublish）
//
// 注意：--watch 只负责「编译」，不会拉起调试会话。要调试请在 VS Code 中按 F5，
// 并选择 "Run Extension (watch)" 启动项（它会自动先跑本 watch 任务）。
const esbuild = require('esbuild');
const path = require('path');
const fs = require('fs');

// 上层 pptx-parser 仓库根目录（本文件位于 <repo>/vscode/）
const PARSER_ROOT = path.resolve(__dirname, '..');

/**
 * 直接定位上层 pptx-parser 源码/资源，替代原来的 npm 包解析。
 * - 入口源码：../src/index.ts
 * - 图表渲染器：../examples/chart-lib/chart-renderer.js
 */
function resolveParserSource(rel) {
  const fromParent = path.resolve(PARSER_ROOT, rel);
  if (!fs.existsSync(fromParent)) {
    throw new Error(
      `Cannot resolve pptx-parser source "${rel}" at "${fromParent}". ` +
      `Make sure the vscode extension lives inside the pptx-parser repository.`
    );
  }
  return fromParent;
}

const watch = process.argv.includes('--watch');
const production = process.argv.includes('--production');

/**
 * Webview（浏览器）bundle 会内联 parser 源码。源码里的 tXml.parseStream 持有一个
 * 间接的 `require('fs')`（仅在从文件路径解析时走到，浏览器路径永不触发）。Stub 掉
 * `fs` 让打包成功，避免把 Node 内置模块带进浏览器产物。
 */
const browserFsStub = {
  name: 'browser-fs-stub',
  setup(build) {
    build.onResolve({ filter: /^fs$/ }, () => ({ path: 'fs', namespace: 'browser-fs-stub' }));
    build.onLoad({ filter: /.*/, namespace: 'browser-fs-stub' }, () => ({
      contents:
        'export default { createReadStream(){}, readFile(){}, readFileSync(){return "";}, ' +
        'writeFileSync(){}, existsSync(){return false;}, statSync(){return {};}, promises:{} };',
      loader: 'js'
    }));
  }
};

// 两个 bundle 都需要把 parser 的 `import css from './css/pptxjs.css'` 当作字符串内联。
// rollup 用 css-as-string 插件实现，esbuild 用 loader 的 text 加载器等价处理。
const cssTextLoader = { '.css': 'text' };

// 把对 @fefeding/ppt-parser 的引用统一改指向上层源码（含 chart-renderer 子路径）。
const parserAlias = {
  '@fefeding/ppt-parser': resolveParserSource('src/index.ts'),
  '@fefeding/ppt-parser/chart-renderer': resolveParserSource('examples/chart-lib/chart-renderer.js')
};

async function buildWithContext() {
  const base = {
    bundle: true,
    sourcemap: !production,
    minify: production,
    logLevel: 'info',
    legalComments: 'none',
    alias: parserAlias,
    loader: cssTextLoader,
    // parser 源码从 ../src 解析 jszip/tinycolor2 时会向上找 pptx-parser/node_modules。
    // 把本插件的 node_modules 也纳入搜索路径，使 CI 只需安装 vscode 目录依赖即可构建。
    nodePaths: [path.join(__dirname, 'node_modules')]
  };

  const extCtx = await esbuild.context({
    ...base,
    entryPoints: [path.join(__dirname, 'src', 'extension.ts')],
    outfile: path.join(__dirname, 'dist', 'extension.js'),
    external: ['vscode'],
    format: 'cjs',
    platform: 'node',
    target: 'node16'
  });

  const webCtx = await esbuild.context({
    ...base,
    entryPoints: [path.join(__dirname, 'src', 'webview', 'main.ts')],
    outfile: path.join(__dirname, 'dist', 'webview.js'),
    format: 'iife',
    platform: 'browser',
    target: 'es2020',
    plugins: [browserFsStub]
  });

  if (watch) {
    await Promise.all([extCtx.watch(), webCtx.watch()]);
    // 该行同时充当 .vscode/tasks.json 里 background 任务的 beginsPattern
    console.log('[esbuild] watching for changes… (press Ctrl+C to stop)');
    console.log('[esbuild] 编译已持续运行。调试：按 F5 选择 "Run Extension (watch)"；');
    console.log('[esbuild] 改代码后在扩展开发宿主窗口按 Ctrl/Cmd+R 刷新即可生效。');
    return;
  }

  await Promise.all([extCtx.rebuild(), webCtx.rebuild()]);
  await extCtx.dispose();
  await webCtx.dispose();
  console.log('[esbuild] build complete.');
}

buildWithContext().catch((e) => {
  console.error(e);
  process.exit(1);
});
