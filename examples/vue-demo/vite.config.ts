import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { resolve } from 'path'

export default defineConfig({
  plugins: [vue()],
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
