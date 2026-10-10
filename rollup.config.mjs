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
  // 打包类型声明文件：生成完整的 .d.ts（自包含，内联 XmlNode/WarpObject 等类型）
  {
    input: 'src/index.ts',
    output: [{ file: pkg.types, format: 'es' }],
    plugins: [cssAsString(), dts()],
    external: [...Object.keys(pkg.dependencies)]
  }
];