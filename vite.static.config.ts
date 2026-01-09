import { defineConfig } from 'vite'

export default defineConfig({
  root: 'src',
  publicDir: 'js',
  server: {
    port: 3001,
    open: true,
  },
  build: {
    outDir: '../dist-static',
    emptyOutDir: true,
  }
})
