<template>
  <div id="app">
    <header class="header">
      <h1>PPTX Parser Demo</h1>
      <p>上传PPTX文件查看HTML渲染结果</p>
    </header>

    <main class="main">
      <div class="upload-section">
        <label class="upload-label">
          <input type="file" accept=".pptx" @change="handleFileUpload" :disabled="loading" />
          <span v-if="!loading">点击选择 PPTX 文件</span>
          <span v-else>解析中... {{ progress }}%</span>
        </label>
      </div>

      <div v-if="error" class="error-message">
        {{ error }}
      </div>

      <div v-if="slides.length > 0" class="preview-section">
        <div class="info-bar">
          <span>共 {{ slides.length }} 张幻灯片</span>
          <button v-if="slideSize" @click="toggleFullscreen" class="fullscreen-btn">全屏</button>
        </div>
        <div class="slide-viewer" ref="slideViewer">
          <div class="slide-container">
            <div class="slides-wrapper">
              <div v-for="(slide, index) in slides" :key="index" class="slide-wrapper">
                <div v-html="slide.html"></div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </main>
  </div>
</template>

<script setup lang="ts">
import { ref, shallowRef, nextTick } from 'vue'
import { pptxToHtml } from '@fefeding/ppt-parser'
import '/Users/jiamao/project/github/pptx-parser/src/lib/jszip.min.js'

interface Slide {
  html: string
  slideNum: number
  fileName: string
}

interface PPTXResult {
  slides: Slide[]
  slideSize?: {
    width: number
    height: number
  }
  styles: {
    global: string
  }
  metadata: Record<string, any>
  charts: any[]
}

const loading = ref(false)
const error = ref('')
const slides = shallowRef<Slide[]>([])
const slideSize = ref<{ width: number; height: number } | null>(null)
const progress = ref(0)
const slideViewer = ref<HTMLElement | null>(null)

async function handleFileUpload(event: Event) {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (!file) return

  // 验证文件类型
  if (file.type !== 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
    error.value = '请选择有效的 PPTX 文件'
    return
  }

  loading.value = true
  error.value = ''
  slides.value = []
  slideSize.value = null
  progress.value = 0

  try {
    // 读取文件为 ArrayBuffer
    const fileData = await file.arrayBuffer()

    // 使用新版本的 API 解析 PPTX
    const result: PPTXResult = await pptxToHtml(fileData, {
      mediaProcess: true,
      themeProcess: true,
      callbacks: {
        onProgress: (percent: number) => {
          progress.value = percent
        }
      }
    })

    // 保存结果
    slides.value = result.slides || []
    slideSize.value = {width: result.slideSize?.width || 0, height: result.slideSize?.height || 0}

    // 等待 DOM 更新后注入全局样式
    await nextTick()
    if (result.styles && result.styles.global) {
      applyGlobalStyles(result.styles.global)
    }

  } catch (e) {
    error.value = e instanceof Error ? e.message : '解析失败'
    console.error('PPTX 解析失败:', e)
  } finally {
    loading.value = false
    // 清空 input 允许重复上传同一文件
    target.value = ''
  }
}

function applyGlobalStyles(css: string) {
  // 查找或创建全局样式容器
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
    viewer.requestFullscreen().catch((err: Error) => {
      console.error('全屏失败:', err)
    })
  }
}
</script>

<style>
@import '/Users/jiamao/project/github/pptx-parser/src/css/pptxjs.css';
</style>

<style scoped>
#app {
  min-height: 100vh;
  background: #f5f5f5;
}

.header {
  background: #4f46e5;
  color: white;
  padding: 1.5rem;
  text-align: center;
}

.header h1 {
  margin: 0 0 0.5rem 0;
  font-size: 1.5rem;
}

.header p {
  margin: 0;
  opacity: 0.9;
  font-size: 1rem;
}

.main {
  margin: 0 auto;
  padding: 1rem;
}

.upload-section {
  margin-bottom: 2rem;
  display: flex;
  justify-content: center;
}

.upload-label {
  display: inline-block;
  padding: 1rem 2rem;
  background: white;
  border: 2px dashed #4f46e5;
  border-radius: 8px;
  cursor: pointer;
  transition: all 0.3s;
  color: #4f46e5;
  font-weight: 600;
}

.upload-label:hover {
  background: #4f46e5;
  color: white;
}

.upload-label input {
  display: none;
}

.error-message {
  background: #fee;
  color: #c33;
  padding: 1rem;
  border-radius: 8px;
  margin-bottom: 1rem;
  text-align: center;
}

.preview-section {
  background: white;
  border-radius: 8px;
  padding: 1rem;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
}

.preview-section h3 {
  margin: 0 0 1rem 0;
  color: #333;
}

.slide-viewer {
  background: #f8f8f8;
  padding: 1rem;
  border-radius: 4px;
  overflow: auto;
  min-height: 400px;
  display: flex;
  justify-content: center;
  align-items: flex-start;
}

.slide-container {
  background: white;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
  border-radius: 4px;
}

.slides-wrapper {
  display: flex;
  flex-direction: column;
  gap: 1rem;
}

.slide-wrapper :deep(section.slide) {
  margin: 0;
  overflow: hidden;
}

.info-bar {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 1rem;
  padding: 0.5rem 1rem;
  background: #f5f5f5;
  border-radius: 4px;
}

.fullscreen-btn {
  padding: 0.5rem 1rem;
  background: #4f46e5;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  transition: background 0.3s;
}

.fullscreen-btn:hover {
  background: #4338ca;
}
</style>

<!-- 全局样式 -->
<style>
/* 为 PPTX 生成的元素添加基础样式 */
#app .slide {
  position: relative;
  background: white;
  overflow: hidden;
}

#app .slide * {
  box-sizing: border-box;
}
</style>
