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
  // 打包核心代码：输出 ESM + CJS 双格式，压缩生产版本
  {
    input: 'src/index.ts',
    output: [
      {
        file: pkg.module,
        format: 'es',
        banner,
        sourcemap: true,
        exports: 'named' // ✅ 修复警告核心配置
      },
      {
        file: pkg.main,
        format: 'cjs',
        banner,
        sourcemap: true,
        exports: 'named' // ✅ 修复警告核心配置
      }
    ],
    plugins: [
      nodeResolve(),
      commonjs(),
      typescript({ tsconfig: './tsconfig.json' }),
      terser({ compress: true, mangle: true })
    ],
    external: [...Object.keys(pkg.dependencies)]
  },
  // 打包类型声明文件：生成完整的.d.ts文件
  {
    input: 'src/index.ts',
    output: [{ file: pkg.types, format: 'es' }],
    plugins: [dts()]
  }
];