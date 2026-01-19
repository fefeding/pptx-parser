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
          <span v-else>解析中...</span>
        </label>
      </div>

      <div v-if="error" class="error-message">
        {{ error }}
      </div>

      <div v-if="htmlContent" class="preview-section">
        <h3>PPTX HTML预览：</h3>
        <div class="slide-viewer">
          <div class="slide-container" v-html="htmlContent"></div>
        </div>
      </div>
    </main>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue'
import { pptxToHtml } from '@fefeding/ppt-parser'

const loading = ref(false)
const error = ref('')
const htmlContent = ref('')

    async function handleFileUpload(event: Event) {
      const target = event.target as HTMLInputElement
      const file = target.files?.[0]
      if (!file) return

      loading.value = true
      error.value = ''
      htmlContent.value = ''

      try {
        // 使用本地重构后的 API：第一个参数是 File，第二个参数是配置项
        const result = await pptxToHtml(file, {
          mediaProcess: true, // 对应原来的 parseImages
          onProgress: (percent: number) => {
            // 可以接上 verbose 的逻辑，这里简单打印
            if (percent % 20 === 0) console.log(`解析进度: ${percent}%`)
          }
        })

        // 新 API 返回 { html, css, slides, ... }，我们取 html 字段显示
        htmlContent.value = result.html

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
  justify-content: center;
  align-items: flex-start;
}

.slide-container {
  background: white;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
  border-radius: 4px;
}
</style>
