<template>
  <div class="image-renderer">
    <img v-if="imageUrl" :src="imageUrl" :style="imageStyle" :alt="element.id" />
    <div v-else class="placeholder">图片</div>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import type { PptNode } from 'pptx-parser'

interface Props {
  element: PptNode
}

const props = defineProps<Props>()

const imageUrl = computed(() => {
  if (props.element.type !== 'image') return ''
  return (props.element.content as any)?.url || ''
})

const imageStyle = computed(() => {
  const style: any = {
    width: '100%',
    height: '100%',
    objectFit: 'contain'
  }

  const element = props.element as any
  if (element.style?.objectFit) {
    style.objectFit = element.style.objectFit
  }

  return style
})
</script>

<style scoped>
.image-renderer {
  width: 100%;
  height: 100%;
}

.image-renderer img {
  display: block;
}

.placeholder {
  width: 100%;
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  background: #f0f0f0;
  color: #999;
  border: 1px dashed #ccc;
}
</style>
