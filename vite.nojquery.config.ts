import { defineConfig } from 'vite'
import { resolve } from 'path'

export default defineConfig({
  root: 'src',
  server: {
    port: 3002,
    open: true,
  },
  build: {
    outDir: '../dist-nojquery',
    emptyOutDir: true,
  },
  resolve: {
    alias: {
      '@fefeding/ppt-parser': resolve(__dirname, 'src/index.ts')
    }
  }
})
