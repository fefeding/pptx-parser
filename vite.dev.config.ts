import { defineConfig } from 'vite'
import { resolve } from 'path'

// 调试专用 Vite 配置：以项目根目录为 root，
// 这样 /examples、/dist、/src 下的文件均可直接通过 HTTP 访问。
export default defineConfig({
  root: resolve(__dirname),
  publicDir: false,
  clearScreen: false,
  server: {
    port: 3001,
    // 默认打开 examples 下的调试页面
    open: '/examples/index.html',
    // 允许访问项目内所有文件（examples / dist / src）
    fs: {
      allow: [resolve(__dirname)]
    },
    // 启用 HMR / 文件变更全量刷新
    hmr: true
  },
  build: {
    outDir: 'dist',
    emptyOutDir: false
  }
})
