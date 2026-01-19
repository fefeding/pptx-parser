import { defineConfig } from 'vite'
import { resolve } from 'path'

export default defineConfig({
  root: 'src',
  publicDir: 'js',
  server: {
    port: 3001,
    open: true,
    fs: {
      allow: [
        resolve(__dirname)
      ]
    },
    hmr: {
      overlay: true
    }
  },
  optimizeDeps: {
    exclude: ['examples']
  },
  plugins: [
    {
      name: 'static-examples',
      configureServer(server) {
        server.middlewares.use('/', async (req, res, next) => {
          const filePath = req.url ? req.url.replace(/^\//, '') : ''
          const fullPath = resolve(__dirname, filePath)
          
          try {
            if (fullPath.endsWith('.html') || fullPath.endsWith('.js') || fullPath.endsWith('.css')) {
              const content = await import('fs').then(fs => fs.promises.readFile(fullPath))
              res.setHeader('Content-Type', fullPath.endsWith('.html') ? 'text/html' : 
                                                   fullPath.endsWith('.js') ? 'application/javascript' : 'text/css')
              res.setHeader('charset', 'utf-8')
              res.end(content)
            } else {
              next()
            }
          } catch (error) {
            next()
          }
        })
      }
    }
  ],
  build: {
    outDir: '../dist',
    emptyOutDir: true,
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      name: 'PPTXParser',
      formats: ['es'],
      fileName: (format) => `pptx-parser.${format}.js`
    },
    rollupOptions: {
      output: {
        preserveModules: false,
        exports: 'named'
      }
    }
  }
})
