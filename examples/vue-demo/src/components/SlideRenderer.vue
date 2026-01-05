<template>
  <div class="slide-content">
    <div v-for="element in elements" :key="element.id" class="ppt-element" :style="getElementStyle(element)">
      <TextRenderer v-if="element.type === 'text'" :element="element" />
      <ImageRenderer v-else-if="element.type === 'image'" :element="element" />
      <ShapeRenderer v-else-if="element.type === 'shape'" :element="element" />
      <TableRenderer v-else-if="element.type === 'table'" :element="element" />
      <div v-else class="unsupported">
        {{ element.type }} (暂不支持)
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import type { PptNode } from 'pptx-parser'
import TextRenderer from './TextRenderer.vue'
import ImageRenderer from './ImageRenderer.vue'
import ShapeRenderer from './ShapeRenderer.vue'
import TableRenderer from './TableRenderer.vue'

interface Props {
  slide: any
}

const props = defineProps<Props>()

const elements = computed(() => {
  return props.slide?.elements || []
})

function getElementStyle(element: PptNode) {
  const style: any = {
    position: 'absolute',
    left: `${element.x}px`,
    top: `${element.y}px`,
    width: `${element.width}px`,
    height: `${element.height}px`,
    zIndex: element.zIndex || 1
  }

  if (element.rotation) {
    style.transform = `rotate(${element.rotation}deg)`
  }

  if (element.opacity !== undefined) {
    style.opacity = element.opacity
  }

  return style
}
</script>

<style scoped>
.slide-content {
  position: relative;
  width: 100%;
  height: 100%;
}

.ppt-element {
  overflow: hidden;
}

.unsupported {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 100%;
  background: #f0f0f0;
  color: #999;
  font-size: 12px;
  border: 1px dashed #ccc;
}
</style>
