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
          <span v-else>解析中... ({{ progress }}%)</span>
        </label>
      </div>

      <div v-if="error" class="error-message">
        {{ error }}
      </div>

      <div v-if="result.htmlContent.length > 0" class="preview-section">
        <h3>PPTX HTML预览 ({{ result.htmlContent.length }} 张幻灯片)：</h3>
        <div class="slide-viewer">
          <div
            v-for="(slide, index) in result.htmlContent"
            :key="index"
            class="slide-item"
            v-html="slide.html"
          ></div>
        </div>
      </div>
    </main>
  </div>
</template>

<script setup lang="ts">
import { ref, watch } from 'vue'
import { pptxToHtml } from '@fefeding/ppt-parser'

interface SlideContent {
  type: string
  html: string
  data?: string
}

interface PPTXResult {
  htmlContent: SlideContent[]
  globalCSS: string
  slideWidth: number
  slideHeight: number
  styleTable: Record<string, any>
  error: any
}

const loading = ref(false)
const error = ref('')
const progress = ref(0)
const result = ref<PPTXResult>({
  htmlContent: [],
  globalCSS: '',
  slideWidth: 0,
  slideHeight: 0,
  styleTable: {},
  error: null
})

// 注入全局 CSS 到 document head
let styleElement: HTMLStyleElement | null = null

watch(() => result.value.globalCSS, (newCSS) => {
  // 移除旧的 style 元素
  if (styleElement) {
    document.head.removeChild(styleElement)
    styleElement = null
  }

  // 如果有新 CSS，创建并插入
  if (newCSS) {
    styleElement = document.createElement('style')
    styleElement.textContent = newCSS
    document.head.appendChild(styleElement)
  }
})

async function handleFileUpload(event: Event) {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (!file) return

  loading.value = true
  error.value = ''
  progress.value = 0
  result.value = {
    htmlContent: [],
    globalCSS: '',
    slideWidth: 0,
    slideHeight: 0,
    styleTable: {},
    error: null
  }

  try {
    const data = await pptxToHtml(file, {
      onProgress: (percent) => {
        progress.value = Math.round(percent)
      },
      onComplete: (res) => {
        console.log('PPTX解析完成:', res)
      },
      onError: (err) => {
        console.error('PPTX解析错误:', err)
        error.value = err instanceof Error ? err.message : '解析失败'
      }
    })

    // 获取生成的HTML内容
    if (data.error) {
      error.value = data.error.message || '解析失败'
    } else {
      result.value = data
    }
  } catch (e) {
    error.value = e instanceof Error ? e.message : '解析失败'
    console.error('PPTX解析失败:', e)
  } finally {
    loading.value = false
  }
}
</script>

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
  flex-direction: column;
  align-items: center;
  gap: 2rem;
}

.slide-item {
  background: white;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
  border-radius: 4px;
  margin: 1rem 0;
}
</style>
