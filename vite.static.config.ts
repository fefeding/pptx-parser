import { defineConfig } from 'vite'
import { resolve } from 'path'
import * as fs from 'fs'

export default defineConfig({
  root: 'src',
  publicDir: 'js',
  server: {
    port: 3001,
    open: true,
    fs: {
      // 允许 Vite 服务器访问项目根目录下的所有文件
      allow: [
        resolve(__dirname),
        resolve(__dirname, 'dist'),
        resolve(__dirname, 'examples'),
        resolve(__dirname, 'src')
      ]
    },
    // 配置服务器代理，将 /examples 请求代理到 examples 目录
    proxy: {
      '/examples': {
        target: 'file://' + resolve(__dirname, 'examples'),
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/examples/, '')
      },
      '/dist': {
        target: 'file://' + resolve(__dirname, 'dist'),
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/dist/, '')
      },
      '/src': {
        target: 'file://' + resolve(__dirname, 'src'),
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/src/, '')
      }
    }
  },
  // 配置依赖优化，避免处理 examples 目录
  optimizeDeps: {
    exclude: ['examples']
  },
  // 配置插件，添加对 examples 和 dist 目录的静态资源服务
  plugins: [
    {
      name: 'static-examples',
      configureServer(server) {
        // 拦截对 /examples 的请求，提供静态文件服务
        server.middlewares.use('/examples', async (req, res, next) => {
          // 移除 /examples 前缀
          const filePath = req.url ? req.url.replace(/^\//, '') : ''
          // 构建完整路径
          const fullPath = resolve(__dirname, 'examples', filePath)

          try {
            if (fs.existsSync(fullPath)) {
              // 读取文件内容
              const content = fs.readFileSync(fullPath)
              // 设置适当的 Content-Type
              if (fullPath.endsWith('.html')) {
                res.setHeader('Content-Type', 'text/html; charset=utf-8')
              } else if (fullPath.endsWith('.js')) {
                res.setHeader('Content-Type', 'application/javascript; charset=utf-8')
              } else if (fullPath.endsWith('.css')) {
                res.setHeader('Content-Type', 'text/css; charset=utf-8')
              }
              // 发送响应
              res.end(content)
            } else {
              // 文件不存在，继续到下一个中间件
              next()
            }
          } catch (error) {
            console.error('Error serving file from examples directory:', error)
            next()
          }
        })

        // 拦截对 /dist 的请求，提供静态文件服务
        server.middlewares.use('/dist', async (req, res, next) => {
          // 移除 /dist 前缀
          const filePath = req.url ? req.url.replace(/^\//, '') : ''
          // 构建完整路径
          const fullPath = resolve(__dirname, 'dist', filePath)

          try {
            if (fs.existsSync(fullPath)) {
              // 读取文件内容
              const content = fs.readFileSync(fullPath)
              // 设置适当的 Content-Type
              if (fullPath.endsWith('.js')) {
                res.setHeader('Content-Type', 'application/javascript; charset=utf-8')
              } else if (fullPath.endsWith('.js.map')) {
                res.setHeader('Content-Type', 'application/json; charset=utf-8')
              } else if (fullPath.endsWith('.d.ts')) {
                res.setHeader('Content-Type', 'text/typescript; charset=utf-8')
              }
              // 发送响应
              res.end(content)
            } else {
              // 文件不存在，继续到下一个中间件
              next()
            }
          } catch (error) {
            console.error('Error serving file from dist directory:', error)
            next()
          }
        }),


        // 拦截对 /dist 的请求，提供静态文件服务
        server.middlewares.use('/src', async (req, res, next) => {
          // 移除 /dist 前缀
          const filePath = req.url ? req.url.replace(/^\//, '') : ''
          // 构建完整路径
          const fullPath = resolve(__dirname, 'src', filePath)

          try {
            if (fs.existsSync(fullPath)) {
              // 读取文件内容
              const content = fs.readFileSync(fullPath)
              // 设置适当的 Content-Type
              if (fullPath.endsWith('.js')) {
                res.setHeader('Content-Type', 'application/javascript; charset=utf-8')
              } else if (fullPath.endsWith('.js.map')) {
                res.setHeader('Content-Type', 'application/json; charset=utf-8')
              } else if (fullPath.endsWith('.d.ts')) {
                res.setHeader('Content-Type', 'text/typescript; charset=utf-8')
              }
              // 发送响应
              res.end(content)
            } else {
              // 文件不存在，继续到下一个中间件
              next()
            }
          } catch (error) {
            console.error('Error serving file from dist directory:', error)
            next()
          }
        })
      }
    }
  ],
  build: {
    outDir: '../dist',
    emptyOutDir: true,
  }
})
