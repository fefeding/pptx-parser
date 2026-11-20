import { nodeResolve } from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';
import dts from 'rollup-plugin-dts';
import fs from 'fs';
import path from 'path';

const pkg = JSON.parse(fs.readFileSync(path.resolve('./package.json'), 'utf-8'));

const banner = `/**
 * ${pkg.name} v${pkg.version}
 * ${pkg.description}
 * MIT License
 */`;

// 把 .css 以内联字符串形式打包，使 src/css/pptxjs.css 成为基础样式的唯一来源
// （避免消费方各自复制一份样式文件）
function cssAsString() {
  return {
    name: 'css-as-string',
    transform(code, id) {
      if (!id.endsWith('.css')) return null;
      return { code: `export default ${JSON.stringify(code)};`, map: { mappings: '' } };
    }
  };
}

// 浏览器转换产物需要内联 pdf-lib 的字体 JSON，而项目未安装 @rollup/plugin-json
function jsonAsEsm() {
  return {
    name: 'json-as-esm',
    transform(code, id) {
      if (!id.endsWith('.json')) return null;
      return { code: `export default ${code};`, map: { mappings: '' } };
    }
  };
}

export default [
  // 打包核心代码：输出 ESM + CJS 双格式，不压缩（用于 Node.js 开发）
  {
    input: 'src/index.ts',
    output: [
      {
        file: './dist/ppt-parser.esm.js',
        format: 'es',
        banner,
        sourcemap: true,
        exports: 'named'
      },
      {
        file: './dist/ppt-parser.cjs',
        format: 'cjs',
        banner,
        sourcemap: true,
        exports: 'named'
      }
    ],
    plugins: [
      nodeResolve({ extensions: ['.ts', '.js', '.json'] }),
      commonjs(),
      typescript({ tsconfig: './tsconfig.json', compilerOptions: { checkJs: false, noEmitOnError: false } }),
      cssAsString()
    ],
    external: [...Object.keys(pkg.dependencies)],
    treeshake: false
  },
  // 打包浏览器版本：输出 ESM 格式（非压缩版本，包含所有依赖）
  {
    input: 'src/index.ts',
    output: {
      file: './dist/ppt-parser.browser.js',
      format: 'es',
      banner,
      sourcemap: true,
      exports: 'named'
    },
    plugins: [
      nodeResolve({
        browser: true,
        preferBuiltins: false,
        extensions: ['.ts', '.js', '.json']
      }),
      commonjs({
        // 将 CJS 模块转换为 ESM
        include: /node_modules/
      }),
      typescript({ tsconfig: './tsconfig.json', compilerOptions: { checkJs: false, noEmitOnError: false } }),
      cssAsString()
    ],
    treeshake: false,
    // 不标记 dependencies 为 external，让它们被打包进去
    // external: []
  },
  // 格式转换子路径（@fefeding/ppt-parser/convert）：插件化转换器注册表。
  // 通过 alias 把 convert 内部对主库的相对自引用（../.. / ../../..）改指向包名，
  // 并标记主包为 external，避免把 PPTX 解析内核重复打进转换 bundle。
  {
    input: 'src/convert/index.ts',
    output: [
      { file: './dist/convert.esm.js', format: 'es', banner, sourcemap: true, exports: 'named' },
      { file: './dist/convert.cjs', format: 'cjs', banner, sourcemap: true, exports: 'named' }
    ],
    plugins: [
      {
        name: 'externalize-core-self-import',
        resolveId(source, importer) {
          if (!importer || !importer.includes('src/convert/')) return null;
          if (/(^|\/)src\/convert\//.test(importer) && ['..', '../..', '../../..'].includes(source)) {
            return { id: '@fefeding/ppt-parser', external: true };
          }
          return null;
        }
      },
      nodeResolve({ extensions: ['.ts', '.js', '.json'] }),
      commonjs(),
      typescript({ tsconfig: './tsconfig.json', compilerOptions: { checkJs: false, noEmitOnError: false } }),
      cssAsString()
    ],
    external: [...Object.keys(pkg.dependencies), '@fefeding/ppt-parser'],
    treeshake: false
  },
  // 格式转换子路径的浏览器版本：第三方依赖（docx / pdf-lib / marked 等）全部内联，
  // 但核心解析内核复用 ppt-parser.browser.js，避免把同一份内核重复打进浏览器产物。
  {
    input: 'src/convert/index.ts',
    output: {
      file: './dist/convert.browser.js',
      format: 'es',
      banner,
      sourcemap: true,
      exports: 'named',
      inlineDynamicImports: true
    },
    plugins: [
      {
        name: 'externalize-core-self-import-browser',
        resolveId(source, importer) {
          if (!importer || !importer.includes('src/convert/')) return null;
          if (/(^|\/)src\/convert\//.test(importer) && ['..', '../..', '../../..'].includes(source)) {
            return { id: './ppt-parser.browser.js', external: true };
          }
          return null;
        }
      },
      nodeResolve({ browser: true, preferBuiltins: false, extensions: ['.ts', '.js', '.json'] }),
      commonjs({ include: /node_modules/ }),
      typescript({ tsconfig: './tsconfig.json', compilerOptions: { checkJs: false, noEmitOnError: false } }),
      cssAsString(),
      jsonAsEsm()
    ],
    // 浏览器端不会进入 Node 专属的系统字体探测分支，这里的 node 内置模块只需保留在产物里
    external: ['node:fs/promises', 'fs'],
    treeshake: false
  },
  // 转换子路径类型声明
  {
    input: 'src/convert/index.ts',
    output: [{ file: './dist/convert.d.ts', format: 'es' }],
    plugins: [cssAsString(), dts()],
    external: [...Object.keys(pkg.dependencies), '@fefeding/ppt-parser']
  },
  // 打包类型声明文件：生成完整的 .d.ts（自包含，内联 XmlNode/WarpObject 等类型）
  {
    input: 'src/index.ts',
    output: [{ file: pkg.types, format: 'es' }],
    plugins: [cssAsString(), dts()],
    external: [...Object.keys(pkg.dependencies)]
  }
];