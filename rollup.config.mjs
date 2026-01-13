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
    input: 'src/index.ts',
    output: [
      {
        file: pkg.module,
        format: 'es',
        banner,
        sourcemap: true,
        exports: 'named'
      },
      {
        file: pkg.main,
        format: 'cjs',
        banner,
        sourcemap: true,
        exports: 'named'
      }
    ],
    plugins: [
      nodeResolve(),
      commonjs(),
      typescript({ tsconfig: './tsconfig.json', compilerOptions: { checkJs: false, noEmitOnError: false } })
    ],
    external: [...Object.keys(pkg.dependencies)]
  },
  // 打包浏览器版本：输出 ESM 格式（压缩版本，包含所有依赖）
  {
    input: 'src/index.ts',
    output: {
      file: './dist/ppt-parser.browser.min.js',
      format: 'es',
      banner,
      sourcemap: true,
      exports: 'named'
    },
    plugins: [
      nodeResolve({
        browser: true,
        preferBuiltins: false
      }),
      commonjs(),
      typescript({ tsconfig: './tsconfig.json', compilerOptions: { checkJs: false, noEmitOnError: false } }),
      terser({
        compress: true,
        mangle: true,
        format: {
          comments: false
        }
      })
    ]
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
        preferBuiltins: false
      }),
      commonjs(),
      typescript({ tsconfig: './tsconfig.json', compilerOptions: { checkJs: false, noEmitOnError: false } })
    ]
  },
  // 打包类型声明文件：生成完整的.d.ts文件
  {
    input: 'src/index.ts',
    output: [{ file: pkg.types, format: 'es' }],
    plugins: [dts()]
  }
];