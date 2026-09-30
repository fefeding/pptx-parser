// 一体化调试服务器：
//   1. rollup watch 监听 src 改动，自动重新打包 dist/ppt-parser.browser.js
//   2. Vite 静态服务器提供 examples/dist/src 的访问，并默认打开 examples/index.html
//   源码或 examples 下任意文件改动后，浏览器自动刷新（热更新）。
// 单一进程，Ctrl+C 同时关闭 rollup watch 与 Vite。
import { fileURLToPath } from 'node:url'
import { dirname, resolve } from 'node:path'
import { watch as rollupWatch } from 'rollup'
import { createServer } from 'vite'

const __dirname = dirname(fileURLToPath(import.meta.url))
const root = __dirname

const rollupConfig = (await import('./rollup.config.mjs')).default

let viteStarted = false
let rollupReady = false

async function startVite() {
  if (viteStarted) return
  viteStarted = true
  const server = await createServer({
    configFile: resolve(root, 'vite.dev.config.ts'),
    root
  })
  await server.listen()
  server.printUrls()

  const shutdown = async () => {
    console.log('\n[dev] Shutting down...')
    await server.close()
    watcher.close()
    process.exit(0)
  }
  process.on('SIGINT', shutdown)
  process.on('SIGTERM', shutdown)
}

const watcher = rollupWatch(rollupConfig)
watcher.on('event', (event) => {
  if (event.code === 'END') {
    // 首次构建完成后启动 Vite，确保 dist 已就绪
    if (!rollupReady) {
      rollupReady = true
      startVite()
    }
  } else if (event.code === 'ERROR') {
    console.error('[rollup] build error:', event.error)
    if (!rollupReady) {
      rollupReady = true
      startVite()
    }
  }
})
