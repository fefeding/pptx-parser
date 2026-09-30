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

export default [
  // 打包核心代码：输出 ESM + CJS 双格式，不压缩（用于 Node.js 开发）
  {
    input: 'src/js/index.ts',
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
      typescript({ tsconfig: './tsconfig.json', compilerOptions: { checkJs: false, noEmitOnError: false } })
    ],
    external: [...Object.keys(pkg.dependencies)]
  },
  // 打包浏览器版本：输出 ESM 格式（非压缩版本，包含所有依赖）
  {
    input: 'src/js/index.ts',
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
      typescript({ tsconfig: './tsconfig.json', compilerOptions: { checkJs: false, noEmitOnError: false } })
    ],
    // 不标记 dependencies 为 external，让它们被打包进去
    // external: []
  },
  // 打包类型声明文件：生成完整的 .d.ts（自包含，内联 XmlNode/WarpObject 等类型）
  {
    input: 'src/js/index.ts',
    output: [{ file: pkg.types, format: 'es' }],
    plugins: [dts()],
    external: [...Object.keys(pkg.dependencies)]
  }
];