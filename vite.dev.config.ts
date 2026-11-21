import { defineConfig } from 'vite'
import { resolve } from 'path'

// 调试专用 Vite 配置：以项目根目录为 root，
// 这样 /examples、/dist、/src 下的文件均可直接通过 HTTP 访问。
// 中文字体代理：浏览器 fetch Google Fonts 会得到 WOFF2，而 pdf-lib 的 fontkit
// 子集化 WOFF2 会丢失字形轮廓（文字能提取但不渲染）。Node fetch 用非浏览器 UA
// 拿到的是 TTF，这里做代理让 examples demo 开箱即用地拿到可用字体。
function cjkFontProxy() {
  return {
    name: 'cjk-font-proxy',
    configureServer(server: { middlewares: { use: (path: string, fn: (req: any, res: any, next: () => void) => void) => void } }) {
      server.middlewares.use('/api/cjk-font', async (req, res) => {
        const url = new URL(req.url || '', 'http://localhost');
        const chars = url.searchParams.get('chars') || '';
        res.setHeader('Content-Type', 'font/ttf');
        try {
          const cssUrl = 'https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400&text=' + encodeURIComponent(chars);
          const css = await (await fetch(cssUrl)).text();
          const m = css.match(/url\((https:[^)]+)\)/);
          if (!m) throw new Error('no font url in CSS');
          const fontBytes = new Uint8Array(await (await fetch(m[1])).arrayBuffer());
          res.end(Buffer.from(fontBytes));
        } catch (err) {
          res.statusCode = 502;
          res.end(String(err));
        }
      });
    }
  };
}

export default defineConfig({
  root: resolve(__dirname),
  publicDir: false,
  clearScreen: false,
  plugins: [cjkFontProxy()],
  server: {
    port: 3001,
    // 默认打开 examples 下的调试页面
    open: '/examples/index.html',
    // 允许访问项目内所有文件（examples / dist / src）
    fs: {
      allow: [resolve(__dirname)]
    },
    // 启用 HMR / 文件变更全量刷新
    hmr: true,
    watch: {
      // dist 由 rollup watch 写入：写入过程中文件是半截的，
      // 这里不让 Vite 直接监听（否则会读到不完整产物报 Pre-transform error），
      // 改由 dev-server.mjs 在每次构建结束后统一触发一次整页刷新。
      ignored: ['**/dist/**']
    }
  },
  build: {
    outDir: 'dist',
    emptyOutDir: false
  }
})
