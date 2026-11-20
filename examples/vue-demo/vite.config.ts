import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { resolve } from 'path'

export default defineConfig({
  plugins: [
    vue(),
    {
      // 库源码把基础 CSS 作为字符串导入，Vite 构建需 ?inline 才提供 default 导出；
      // dev 下同样可用，保留对上层 src 源码的热加载。
      name: 'pptx-parser-css-inline',
      enforce: 'pre',
      transform(code, id) {
        if (id.endsWith('/src/index.ts') && code.includes("from './css/pptxjs.css'")) {
          return { code: code.replace("from './css/pptxjs.css'", "from './css/pptxjs.css?inline'"), map: null }
        }
      }
    }
  ],
  resolve: {
    alias: {
      '@': resolve(__dirname, 'src'),
      '@fefeding/ppt-parser': resolve(__dirname, '../../src/index.ts')
    }
  },
  optimizeDeps: {
    exclude: ['@fefeding/ppt-parser']
  },
  envDir: process.cwd(),
  server: {
    port: 3000,
    open: true,
    watch: {
      // 监听上层库的源码变化
      ignored: ['!../../src/**']
    }
  }
})
