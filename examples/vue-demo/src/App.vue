<template>
  <div id="app">
    <header class="header">
      <h1>PPTX Parser Demo</h1>
      <p>上传PPTX文件查看解析结果（使用toHTML渲染）</p>
    </header>

    <main class="main">
      <div class="upload-section">
        <label class="upload-label">
          <input type="file" accept=".pptx" @change="handleFileUpload" :disabled="loading" />
          <span v-if="!loading">点击选择 PPTX 文件</span>
          <span v-else>解析中...</span>
        </label>
      </div>

      <div v-if="error" class="error-message">
        {{ error }}
      </div>

      <div v-if="parsedData" class="preview-section">
        <div class="controls">
          <button @click="currentSlideIndex = Math.max(0, currentSlideIndex - 1)" :disabled="currentSlideIndex === 0">
            ← 上一页
          </button>
          <span class="slide-counter">{{ currentSlideIndex + 1 }} / {{ parsedData.slides.length }}</span>
          <button @click="currentSlideIndex = Math.min(parsedData.slides.length - 1, currentSlideIndex + 1)" :disabled="currentSlideIndex === parsedData.slides.length - 1">
            下一页 →
          </button>
          <button @click="exportHTML" style="margin-left: 1rem; background: #10b981;">
            导出HTML
          </button>
        </div>

        <div class="slide-thumbnails">
          <div
            v-for="(slide, index) in parsedData.slides"
            :key="slide.id"
            class="thumbnail"
            :class="{ active: index === currentSlideIndex }"
            @click="currentSlideIndex = index"
          >
            <span>幻灯片 {{ index + 1 }}</span>
          </div>
        </div>

        <div class="slide-viewer">
          <div class="slide" :style="slideStyle" v-html="currentSlideHTML"></div>
        </div>

        <div class="raw-data">
          <details>
            <summary>查看原始数据 (JSON)</summary>
            <pre>{{ JSON.stringify(currentSlide, null, 2) }}</pre>
          </details>
        </div>
      </div>
    </main>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'
import { parsePptx } from 'pptx-parser'
import type { PptxParseResult } from 'pptx-parser'

const loading = ref(false)
const error = ref('')
const parsedData = ref<PptxParseResult | null>(null)
const currentSlideIndex = ref(0)

const currentSlide = computed(() => {
  if (!parsedData.value || !parsedData.value.slides[currentSlideIndex.value]) {
    return null
  }
  return parsedData.value.slides[currentSlideIndex.value]
})

const currentSlideHTML = computed(() => {
  if (!parsedData.value || !currentSlide.value) return ''

  // 直接使用元素实例的toHTML方法渲染
  const slide = currentSlide.value
  const containerStyle = [
    `width: 100%`,
    `height: 100%`,
    `position: relative`,
    `background-color: ${slide.background || '#ffffff'}`,
    `overflow: hidden`
  ].join('; ')

  const elementsHTML = slide.elements.map((element: any) => {
    // 元素已经是BaseElement实例，直接调用toHTML
    return element.toHTML ? element.toHTML() : ''
  }).join('\n')

  return `<div style="${containerStyle}">
${elementsHTML}
    </div>`
})

const slideStyle = computed(() => {
  if (!parsedData.value) return {}
  const { width, height } = parsedData.value.props
  return {
    width: `${width}px`,
    height: `${height}px`
  }
})

async function handleFileUpload(event: Event) {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (!file) return

  loading.value = true
  error.value = ''

  try {
    const arrayBuffer = await file.arrayBuffer()
    const data = await parsePptx(arrayBuffer, {
      parseImages: true,
      keepRawXml: false,
      verbose: true
    })
    parsedData.value = data
    currentSlideIndex.value = 0
    console.log('解析结果:', data)
  } catch (err) {
    error.value = err instanceof Error ? err.message : '解析失败，请检查文件格式'
    console.error('Parse error:', err)
  } finally {
    loading.value = false
  }
}

function exportHTML() {
  if (!parsedData.value) return

  const slidesHTML = parsedData.value.slides.map((slide: any) => {
    const containerStyle = [
      `width: 100%`,
      `height: 100%`,
      `position: relative`,
      `background-color: ${slide.background || '#ffffff'}`,
      `overflow: hidden`
    ].join('; ')

    const elementsHTML = slide.elements.map((element: any) => {
      return element.toHTML ? element.toHTML() : ''
    }).join('\n')

    return `<div class="ppt-slide" style="${containerStyle}" data-slide-id="${slide.id}">
${elementsHTML}
    </div>`
  }).join('\n\n')

  const { width, height } = parsedData.value.props
  const htmlContent = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${parsedData.value.title || 'PPTX Presentation'}</title>
  <style>
    body {
      margin: 0;
      padding: 20px;
      background: #f5f5f5;
      font-family: Arial, sans-serif;
    }
    .ppt-container {
      max-width: ${width}px;
      margin: 0 auto;
    }
    .ppt-slide {
      width: ${width}px;
      height: ${height}px;
      background: white;
      margin: 20px auto;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
  </style>
</head>
<body>
  <div class="ppt-container">
    <h1 style="text-align: center; margin-bottom: 30px;">${parsedData.value.title || 'PPTX Presentation'}</h1>
${slidesHTML}
  </div>
</body>
</html>`

  const blob = new Blob([htmlContent], { type: 'text/html' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `${parsedData.value.title || 'presentation'}.html`
  a.click()
  URL.revokeObjectURL(url)
}
</script>

<style scoped>
#app {
  min-height: 100vh;
  background: #f5f5f5;
}

.header {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  padding: 2rem;
  text-align: center;
}

.header h1 {
  margin: 0 0 0.5rem 0;
  font-size: 2rem;
}

.header p {
  margin: 0;
  opacity: 0.9;
}

.main {
  max-width: 1400px;
  margin: 0 auto;
  padding: 2rem;
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
  border: 2px dashed #667eea;
  border-radius: 8px;
  cursor: pointer;
  transition: all 0.3s;
  color: #667eea;
  font-weight: 600;
}

.upload-label:hover {
  background: #667eea;
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
  margin-bottom: 2rem;
  text-align: center;
}

.preview-section {
  background: white;
  border-radius: 12px;
  padding: 2rem;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
}

.controls {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 1rem;
  margin-bottom: 1.5rem;
}

.controls button {
  padding: 0.5rem 1.5rem;
  background: #667eea;
  color: white;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  transition: background 0.3s;
}

.controls button:hover:not(:disabled) {
  background: #5568d3;
}

.controls button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.slide-counter {
  font-weight: 600;
  color: #333;
  min-width: 100px;
  text-align: center;
}

.slide-thumbnails {
  display: flex;
  gap: 0.5rem;
  margin-bottom: 1.5rem;
  overflow-x: auto;
  padding: 0.5rem;
}

.thumbnail {
  flex-shrink: 0;
  padding: 0.5rem 1rem;
  background: #f0f0f0;
  border-radius: 4px;
  cursor: pointer;
  font-size: 0.875rem;
  transition: all 0.2s;
}

.thumbnail:hover {
  background: #e0e0e0;
}

.thumbnail.active {
  background: #667eea;
  color: white;
}

.slide-viewer {
  display: flex;
  justify-content: center;
  margin-bottom: 2rem;
  background: #e0e0e0;
  padding: 2rem;
  border-radius: 8px;
  overflow: auto;
}

.slide {
  background: white;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
  position: relative;
  overflow: hidden;
}

.raw-data {
  margin-top: 2rem;
}

.raw-data details {
  background: #f8f8f8;
  padding: 1rem;
  border-radius: 8px;
  cursor: pointer;
}

.raw-data summary {
  font-weight: 600;
  color: #667eea;
}

.raw-data pre {
  margin: 1rem 0 0 0;
  padding: 1rem;
  background: white;
  border-radius: 4px;
  overflow-x: auto;
  max-height: 500px;
  font-size: 0.75rem;
}
</style>
