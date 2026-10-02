<template>
  <div id="app">
    <header class="header">
      <div class="header-inner">
        <div class="logo">
          <svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <path d="M14 2v6h6M9 13h6M9 17h6"/>
          </svg>
          <div>
            <h1>PPTX Parser</h1>
            <p>上传 PPTX 文件，查看渲染结果与解析 JSON</p>
          </div>
        </div>

        <label class="upload-btn">
          <input type="file" accept=".pptx" @change="onFileChange" :disabled="loading" />
          <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/>
          </svg>
          <span>选择文件</span>
        </label>
      </div>
    </header>

    <main class="main">
      <!-- 拖拽 / 空状态 -->
      <div
        v-if="!hasResult"
        class="dropzone"
        :class="{ dragging: isDragging }"
        @dragover.prevent="isDragging = true"
        @dragleave.prevent="isDragging = false"
        @drop.prevent="onDrop"
      >
        <svg viewBox="0 0 24 24" width="44" height="44" fill="none" stroke="currentColor" stroke-width="1.5">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <path d="M17 8l-5-5-5 5M12 3v12"/>
        </svg>
        <p class="drop-title">将 PPTX 文件拖拽到此处</p>
        <p class="drop-hint">或点击右上角「选择文件」按钮上传</p>
      </div>

      <!-- 加载中 -->
      <div v-if="loading && !hasResult" class="loading-card">
        <div class="spinner"></div>
        <p>正在解析中… {{ progress }}%</p>
        <div class="progress-bar"><div class="progress-fill" :style="{ width: progress + '%' }"></div></div>
      </div>

      <div v-if="error" class="error-message">
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2">
          <circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/>
        </svg>
        <span>{{ error }}</span>
      </div>

      <!-- 结果区 -->
      <section v-if="hasResult" class="result">
        <div class="meta-bar">
          <div class="meta-item">
            <span class="meta-label">文件</span>
            <span class="meta-value" :title="fileName">{{ fileName }}</span>
          </div>
          <div class="meta-item">
            <span class="meta-label">大小</span>
            <span class="meta-value">{{ fileSizeText }}</span>
          </div>
          <div class="meta-item">
            <span class="meta-label">幻灯片</span>
            <span class="meta-value">{{ slides.length }} 张</span>
          </div>
          <div class="meta-item" v-if="slideSize">
            <span class="meta-label">尺寸</span>
            <span class="meta-value">{{ slideSize.width }} × {{ slideSize.height }}</span>
          </div>
          <button class="reset-btn" @click="reset">重新上传</button>
        </div>

        <div class="tabs">
          <button class="tab" :class="{ active: activeTab === 'render' }" @click="activeTab = 'render'">
            渲染预览
          </button>
          <button class="tab" :class="{ active: activeTab === 'json' }" @click="activeTab = 'json'">
            JSON 数据
          </button>
        </div>

        <!-- 渲染预览 -->
        <div v-show="activeTab === 'render'" class="tab-panel">
          <div class="info-bar">
            <span>共 {{ slides.length }} 张幻灯片</span>
            <div class="view-controls">
              <button class="zoom-btn" @click="zoomOut" title="缩小">−</button>
              <button class="zoom-btn zoom-label" @click="fitToWidth" title="适应宽度">
                {{ Math.round(scale * 100) }}%
              </button>
              <button class="zoom-btn" @click="zoomIn" title="放大">+</button>
              <button v-if="slideSize" @click="toggleFullscreen" class="fullscreen-btn">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2">
                  <path d="M8 3H5a2 2 0 0 0-2 2v3M16 3h3a2 2 0 0 1 2 2v3M8 21H5a2 2 0 0 1-2-2v-3M16 21h3a2 2 0 0 0 2-2v-3"/>
                </svg>
                全屏
              </button>
            </div>
          </div>
          <div class="slide-viewer" ref="slideViewer">
            <div class="slide-container">
              <div class="slides-wrapper">
                <div
                  v-for="(slide, index) in slides"
                  :key="index"
                  class="slide-scaler"
                  :style="scalerStyle"
                >
                  <div class="slide-inner" v-html="slide.html" :style="innerScaleStyle"></div>
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- JSON 数据 -->
        <div v-show="activeTab === 'json'" class="tab-panel">
          <div class="json-toolbar">
            <div class="slide-chips">
              <button
                class="chip"
                :class="{ active: selectedSlide === 'all' }"
                @click="selectedSlide = 'all'"
              >全部</button>
              <button
                v-for="(_, index) in slides"
                :key="index"
                class="chip"
                :class="{ active: selectedSlide === index }"
                @click="selectedSlide = index"
              >第 {{ index + 1 }} 页</button>
            </div>
            <div class="json-actions">
              <label class="switch">
                <input type="checkbox" v-model="stripLargeFields" />
                <span>精简大字段</span>
              </label>
              <button class="copy-btn" @click="copyJson">
                <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2">
                  <rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
                </svg>
                复制
              </button>
            </div>
          </div>
          <pre class="json-view"><code v-html="highlightedJson"></code></pre>
        </div>
      </section>
    </main>
  </div>
</template>

<script setup lang="ts">
import { ref, shallowRef, computed, nextTick, watch, onMounted, onUnmounted } from 'vue'
import { pptxToHtml, pptxToJson } from '@fefeding/ppt-parser'
import JSZip from 'jszip'
import * as echarts from 'echarts'
// echarts-gl：注册 bar3D / line3D / surface 等真 3D 坐标系，须在 echarts 之后引入
import * as echartsGL from 'echarts-gl'
import { chartRenderer } from '../../chart-lib/chart-renderer.js'

// 设置全局 JSZip 对象，供 ppt-parser 使用
;(window as any).JSZip = JSZip
;(window as any).echarts = echarts
;(window as any).chartRenderer = chartRenderer
// chart-renderer 通过该全局标志判断 3D 能力；UMD 引入时会自动挂载，
// 这里走 ESM 打包路径，需手动暴露
;(window as any)['echarts-gl'] = echartsGL

interface Slide {
  html: string
  slideNum: number
  fileName: string
}

interface PPTXResult {
  slides: Slide[]
  slideSize?: { width: number; height: number }
  styles: { global: string }
  metadata: Record<string, any>
  charts: any[]
}

const loading = ref(false)
const error = ref('')
const slides = shallowRef<Slide[]>([])
const slideSize = ref<{ width: number; height: number } | null>(null)
const progress = ref(0)
const slideViewer = ref<HTMLElement | null>(null)

const hasResult = computed(() => slides.value.length > 0)
const activeTab = ref<'render' | 'json'>('render')

// 幻灯片缩放（适应宽度）
const scale = ref(1)
const scalerStyle = computed(() => ({
  width: `${(slideSize.value?.width || 0) * scale.value}px`,
  height: `${(slideSize.value?.height || 0) * scale.value}px`
}))
const innerScaleStyle = computed(() => ({
  transform: `scale(${scale.value})`,
  transformOrigin: 'top left'
}))

// 根据视口宽度计算适应缩放比（限制 0.1~2 倍）
function fitToWidth() {
  const viewer = slideViewer.value
  if (!viewer || !slideSize.value?.width) {
    scale.value = 1
    return
  }
  const avail = viewer.clientWidth - 48 // 减去 padding (1.5rem * 2)
  const s = avail / slideSize.value.width
  scale.value = Math.max(0.1, Math.min(2, s))
}
function zoomIn() {
  scale.value = Math.min(2, +(scale.value * 1.1).toFixed(3))
}
function zoomOut() {
  scale.value = Math.max(0.1, +(scale.value / 1.1).toFixed(3))
}
function onViewportResize() {
  if (activeTab.value === 'render') fitToWidth()
}

// 切换到渲染页 / 窗口尺寸变化 / 全屏时重新适应宽度
watch(activeTab, (tab) => {
  if (tab === 'render') nextTick(fitToWidth)
})
onMounted(() => {
  window.addEventListener('resize', onViewportResize)
  document.addEventListener('fullscreenchange', onViewportResize)
})
onUnmounted(() => {
  window.removeEventListener('resize', onViewportResize)
  document.removeEventListener('fullscreenchange', onViewportResize)
})

// JSON 视图状态
const jsonResult = ref<any>(null)
const selectedSlide = ref<'all' | number>('all')
const stripLargeFields = ref(true)

// 文件元信息
const fileName = ref('')
const fileSize = ref(0)
const fileSizeText = computed(() => {
  const b = fileSize.value
  if (b < 1024) return `${b} B`
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`
  return `${(b / 1024 / 1024).toFixed(2)} MB`
})

const isDragging = ref(false)

// ---------- 文件处理 ----------
function onFileChange(event: Event) {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (file) handleFile(file)
  target.value = ''
}

function onDrop(event: DragEvent) {
  isDragging.value = false
  const file = event.dataTransfer?.files?.[0]
  if (file) handleFile(file)
}

function handleFile(file: File) {
  const isPptx =
    file.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' ||
    file.name.toLowerCase().endsWith('.pptx')
  if (!isPptx) {
    error.value = '请选择有效的 PPTX 文件'
    return
  }

  loading.value = true
  error.value = ''
  slides.value = []
  slideSize.value = null
  jsonResult.value = null
  progress.value = 0
  fileName.value = file.name
  fileSize.value = file.size
  activeTab.value = 'render'

  file.arrayBuffer().then(runParse).catch((e) => {
    error.value = e instanceof Error ? e.message : '读取文件失败'
    loading.value = false
  })
}

async function runParse(fileData: ArrayBuffer) {
  try {
    // 结构化 JSON（用于展示）
    const json = await pptxToJson(fileData)
    jsonResult.value = json

    // 渲染 HTML
    const result: PPTXResult = await pptxToHtml(fileData, {
      mediaProcess: true,
      themeProcess: true,
      callbacks: {
        onProgress: (percent: number) => {
          progress.value = percent
        }
      }
    })

    slides.value = result.slides || []
    slideSize.value = {
      width: result.slideSize?.width || 0,
      height: result.slideSize?.height || 0
    }

    await nextTick()
    if (result.styles?.global) applyGlobalStyles(result.styles.global)

    // 幻灯片自适应宽度
    await nextTick()
    fitToWidth()

    if (result.charts?.length) {
      await nextTick()
      chartRenderer.renderCharts(result.charts)
    }
  } catch (e) {
    error.value = e instanceof Error ? e.message : '解析失败'
    console.error('PPTX 解析失败:', e)
  } finally {
    loading.value = false
  }
}

function reset() {
  slides.value = []
  slideSize.value = null
  jsonResult.value = null
  error.value = ''
  progress.value = 0
}

function applyGlobalStyles(css: string) {
  let styleEl = document.getElementById('pptx-global-styles')
  if (!styleEl) {
    styleEl = document.createElement('style')
    styleEl.id = 'pptx-global-styles'
    document.head.appendChild(styleEl)
  }
  styleEl.innerHTML = css
}

function toggleFullscreen() {
  const viewer = slideViewer.value
  if (!viewer) return
  if (document.fullscreenElement) {
    document.exitFullscreen()
  } else {
    viewer.requestFullscreen().catch((err: Error) => console.error('全屏失败:', err))
  }
}

// ---------- JSON 展示 ----------
const displayJson = computed(() => {
  if (!jsonResult.value) return ''
  const source =
    selectedSlide.value === 'all'
      ? jsonResult.value
      : jsonResult.value.slides?.[selectedSlide.value]?.data
  const data = stripLargeFields.value ? sanitize(source) : source
  return JSON.stringify(data, null, 2)
})

const highlightedJson = computed(() => syntaxHighlight(displayJson.value))

function sanitize(obj: any, depth = 0): any {
  if (typeof obj === 'string') {
    return obj.length > 300 ? `«大型数据 ${obj.length} 字符，已折叠»` : obj
  }
  if (obj instanceof Array) return obj.map((v) => sanitize(v, depth + 1))
  if (obj && typeof obj === 'object') {
    const out: Record<string, any> = {}
    for (const key of Object.keys(obj)) {
      if (['thumbnail', 'base64', 'dataUrl', 'image'].includes(key)) {
        const v = obj[key]
        out[key] =
          typeof v === 'string' ? `«${key}: ${v.length} 字符，已折叠»` : v
        continue
      }
      out[key] = sanitize(obj[key], depth + 1)
    }
    return out
  }
  return obj
}

function syntaxHighlight(json: string): string {
  const escaped = json
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
  return escaped.replace(
    /("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g,
    (match) => {
      let cls = 'json-number'
      if (/^"/.test(match)) {
        cls = /:$/.test(match) ? 'json-key' : 'json-string'
      } else if (/true|false/.test(match)) {
        cls = 'json-boolean'
      } else if (/null/.test(match)) {
        cls = 'json-null'
      }
      return `<span class="${cls}">${match}</span>`
    }
  )
}

async function copyJson() {
  try {
    await navigator.clipboard.writeText(displayJson.value)
    copied.value = true
    setTimeout(() => (copied.value = false), 1500)
  } catch {
    error.value = '复制失败，请手动选择文本'
  }
}
const copied = ref(false)
</script>

<style>
@import '../../../src/css/pptxjs.css';
</style>

<style scoped>
:root {
  --primary: #4f46e5;
  --primary-dark: #4338ca;
  --bg: #f4f5fb;
  --card: #ffffff;
  --text: #1f2430;
  --muted: #6b7280;
  --border: #e5e7eb;
}

#app {
  min-height: 100vh;
  background: var(--bg);
  color: var(--text);
}

/* Header */
.header {
  background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
  color: #fff;
  box-shadow: 0 4px 20px rgba(79, 70, 229, 0.25);
}
.header-inner {
  max-width: 1200px;
  margin: 0 auto;
  padding: 1.2rem 1.5rem;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1rem;
}
.logo {
  display: flex;
  align-items: center;
  gap: 0.85rem;
}
.logo h1 {
  margin: 0;
  font-size: 1.4rem;
  font-weight: 700;
  letter-spacing: 0.5px;
}
.logo p {
  margin: 0;
  font-size: 0.85rem;
  opacity: 0.85;
}
.upload-btn {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  padding: 0.6rem 1.2rem;
  background: rgba(255, 255, 255, 0.15);
  border: 1px solid rgba(255, 255, 255, 0.4);
  border-radius: 999px;
  color: #fff;
  font-weight: 600;
  cursor: pointer;
  transition: background 0.2s, transform 0.1s;
  white-space: nowrap;
}
.upload-btn:hover {
  background: rgba(255, 255, 255, 0.28);
}
.upload-btn:active {
  transform: scale(0.97);
}
.upload-btn input {
  display: none;
}

/* Main */
.main {
  max-width: 1200px;
  margin: 0 auto;
  padding: 1.5rem;
}

/* Dropzone */
.dropzone {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 0.6rem;
  padding: 4rem 2rem;
  background: var(--card);
  border: 2px dashed var(--border);
  border-radius: 16px;
  color: var(--muted);
  transition: border-color 0.2s, background 0.2s, transform 0.2s;
}
.dropzone.dragging {
  border-color: var(--primary);
  background: #eef2ff;
  transform: scale(1.01);
  color: var(--primary);
}
.drop-title {
  margin: 0;
  font-size: 1.1rem;
  font-weight: 600;
  color: var(--text);
}
.drop-hint {
  margin: 0;
  font-size: 0.9rem;
}

/* Loading */
.loading-card {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 1rem;
  padding: 3rem;
  background: var(--card);
  border-radius: 16px;
  box-shadow: 0 4px 24px rgba(0, 0, 0, 0.05);
}
.spinner {
  width: 42px;
  height: 42px;
  border: 4px solid #e0e7ff;
  border-top-color: var(--primary);
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}
@keyframes spin {
  to { transform: rotate(360deg); }
}
.progress-bar {
  width: 280px;
  height: 8px;
  background: #e5e7eb;
  border-radius: 999px;
  overflow: hidden;
}
.progress-fill {
  height: 100%;
  background: linear-gradient(90deg, #4f46e5, #7c3aed);
  transition: width 0.2s;
}

/* Error */
.error-message {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  background: #fef2f2;
  color: #dc2626;
  border: 1px solid #fecaca;
  padding: 0.9rem 1.2rem;
  border-radius: 12px;
  margin-bottom: 1rem;
}

/* Result */
.result {
  background: var(--card);
  border-radius: 16px;
  box-shadow: 0 4px 24px rgba(0, 0, 0, 0.05);
  overflow: hidden;
}
.meta-bar {
  display: flex;
  align-items: center;
  gap: 1.5rem;
  padding: 1rem 1.5rem;
  border-bottom: 1px solid var(--border);
  flex-wrap: wrap;
}
.meta-item {
  display: flex;
  flex-direction: column;
  min-width: 0;
}
.meta-label {
  font-size: 0.72rem;
  color: var(--muted);
  text-transform: uppercase;
  letter-spacing: 0.5px;
}
.meta-value {
  font-size: 0.95rem;
  font-weight: 600;
  max-width: 220px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.reset-btn {
  margin-left: auto;
  padding: 0.5rem 1rem;
  background: #f3f4f6;
  border: 1px solid var(--border);
  border-radius: 8px;
  color: var(--text);
  cursor: pointer;
  font-weight: 500;
  transition: background 0.2s;
}
.reset-btn:hover {
  background: #e5e7eb;
}

/* Tabs */
.tabs {
  display: flex;
  gap: 0.25rem;
  padding: 0.75rem 1.5rem 0;
  border-bottom: 1px solid var(--border);
}
.tab {
  padding: 0.6rem 1.2rem;
  background: transparent;
  border: none;
  border-bottom: 2px solid transparent;
  color: var(--muted);
  font-weight: 600;
  font-size: 0.95rem;
  cursor: pointer;
  transition: color 0.2s, border-color 0.2s;
}
.tab:hover {
  color: var(--primary);
}
.tab.active {
  color: var(--primary);
  border-bottom-color: var(--primary);
}
.tab-panel {
  padding: 1.5rem;
}

/* Render info bar */
.info-bar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 1rem;
  color: var(--muted);
  font-size: 0.9rem;
}
.fullscreen-btn {
  display: inline-flex;
  align-items: center;
  gap: 0.4rem;
  padding: 0.45rem 0.9rem;
  background: var(--primary);
  color: #fff;
  border: none;
  border-radius: 8px;
  cursor: pointer;
  font-weight: 500;
  transition: background 0.2s;
}
.fullscreen-btn:hover {
  background: var(--primary-dark);
}

/* 缩放控制 */
.view-controls {
  display: flex;
  align-items: center;
  gap: 0.4rem;
}
.zoom-btn {
  min-width: 32px;
  height: 32px;
  padding: 0 0.5rem;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  background: #fff;
  border: 1px solid var(--border);
  border-radius: 8px;
  color: var(--text);
  font-size: 1rem;
  cursor: pointer;
  transition: all 0.15s;
}
.zoom-btn:hover {
  border-color: var(--primary);
  color: var(--primary);
}
.zoom-label {
  font-size: 0.82rem;
  font-weight: 600;
  min-width: 56px;
}

/* Slide viewer */
.slide-viewer {
  background: #f8f8fb;
  padding: 1.5rem;
  border-radius: 12px;
  overflow: auto;
  min-height: 400px;
}
.slide-container {
  width: fit-content;
  margin: 0 auto;
}
.slides-wrapper {
  display: flex;
  flex-direction: column;
  gap: 1.5rem;
}
/* 缩放层：固定为缩放后的尺寸，内部 slide 通过 transform 缩放 */
.slide-scaler {
  position: relative;
  overflow: hidden;
  background: #fff;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
  border-radius: 6px;
}
.slide-inner {
  position: absolute;
  top: 0;
  left: 0;
  transform-origin: top left;
}
.slide-scaler :deep(section.slide) {
  margin: 0;
  overflow: hidden;
}

/* JSON toolbar */
.json-toolbar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 1rem;
  flex-wrap: wrap;
  margin-bottom: 1rem;
}
.slide-chips {
  display: flex;
  flex-wrap: wrap;
  gap: 0.4rem;
}
.chip {
  padding: 0.35rem 0.8rem;
  background: #f3f4f6;
  border: 1px solid var(--border);
  border-radius: 999px;
  color: var(--muted);
  font-size: 0.82rem;
  cursor: pointer;
  transition: all 0.2s;
}
.chip:hover {
  border-color: var(--primary);
  color: var(--primary);
}
.chip.active {
  background: var(--primary);
  border-color: var(--primary);
  color: #fff;
}
.json-actions {
  display: flex;
  align-items: center;
  gap: 1rem;
}
.switch {
  display: inline-flex;
  align-items: center;
  gap: 0.4rem;
  font-size: 0.82rem;
  color: var(--muted);
  cursor: pointer;
}
.copy-btn {
  display: inline-flex;
  align-items: center;
  gap: 0.35rem;
  padding: 0.4rem 0.9rem;
  background: #f3f4f6;
  border: 1px solid var(--border);
  border-radius: 8px;
  color: var(--text);
  font-size: 0.85rem;
  font-weight: 500;
  cursor: pointer;
  transition: background 0.2s;
}
.copy-btn:hover {
  background: #e5e7eb;
}

/* JSON code */
.json-view {
  margin: 0;
  padding: 1.2rem 1.4rem;
  background: #1e1e2e;
  color: #cdd6f4;
  border-radius: 12px;
  font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
  font-size: 0.82rem;
  line-height: 1.6;
  max-height: 70vh;
  overflow: auto;
  white-space: pre;
  tab-size: 2;
}
.json-view .json-key { color: #89b4fa; }
.json-view .json-string { color: #a6e3a1; }
.json-view .json-number { color: #fab387; }
.json-view .json-boolean { color: #f38ba8; }
.json-view .json-null { color: #9399b2; }
</style>

<!-- 全局样式 -->
<style>
#app .slide {
  position: relative;
  background: white;
  overflow: hidden;
}
#app .slide * {
  box-sizing: border-box;
}
</style>
