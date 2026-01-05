<template>
  <div class="shape-renderer" :style="shapeStyle">
    <svg v-if="shapeType === 'rect'" :width="width" :height="height" :viewBox="`0 0 ${width} ${height}`">
      <rect
        :width="width"
        :height="height"
        :rx="borderRadius"
        :ry="borderRadius"
        :fill="fill"
        :stroke="stroke"
        :stroke-width="strokeWidth"
      />
    </svg>
    <svg v-else-if="shapeType === 'circle'" :width="width" :height="height" :viewBox="`0 0 ${width} ${height}`">
      <ellipse
        :cx="width / 2"
        :cy="height / 2"
        :rx="width / 2"
        :ry="height / 2"
        :fill="fill"
        :stroke="stroke"
        :stroke-width="strokeWidth"
      />
    </svg>
    <svg v-else-if="shapeType === 'triangle'" :width="width" :height="height" :viewBox="`0 0 ${width} ${height}`">
      <polygon
        :points="`${width / 2},0 ${width},${height} 0,${height}`"
        :fill="fill"
        :stroke="stroke"
        :stroke-width="strokeWidth"
      />
    </svg>
    <svg v-else-if="shapeType === 'roundedRect'" :width="width" :height="height" :viewBox="`0 0 ${width} ${height}`">
      <rect
        :width="width"
        :height="height"
        :rx="10"
        :ry="10"
        :fill="fill"
        :stroke="stroke"
        :stroke-width="strokeWidth"
      />
    </svg>
    <svg v-else-if="shapeType === 'diamond'" :width="width" :height="height" :viewBox="`0 0 ${width} ${height}`">
      <polygon
        :points="`${width / 2},0 ${width},${height / 2} ${width / 2},${height} 0,${height / 2}`"
        :fill="fill"
        :stroke="stroke"
        :stroke-width="strokeWidth"
      />
    </svg>
    <svg v-else :width="width" :height="height" :viewBox="`0 0 ${width} ${height}`">
      <rect
        :width="width"
        :height="height"
        :fill="fill"
        :stroke="stroke"
        :stroke-width="strokeWidth"
      />
    </svg>

    <div v-if="hasText" class="shape-text" :style="textStyle">
      <TextRenderer :element="{ ...element, type: 'text' as const }" />
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import type { PptNode } from 'pptx-parser'
import TextRenderer from './TextRenderer.vue'

interface Props {
  element: PptNode
}

const props = defineProps<Props>()

const width = computed(() => props.element.width)
const height = computed(() => props.element.height)

const content = computed(() => {
  if (props.element.type !== 'shape') return null
  return props.element.content as any
})

const shapeType = computed(() => {
  return content.value?.shapeType || 'rect'
})

const borderRadius = computed(() => {
  return content.value?.borderRadius || 0
})

const fill = computed(() => {
  const style = props.element.style as any
  return style?.backgroundColor || '#ffffff'
})

const stroke = computed(() => {
  const style = props.element.style as any
  return style?.borderColor || 'none'
})

const strokeWidth = computed(() => {
  const style = props.element.style as any
  return style?.borderWidth || 0
})

const hasText = computed(() => {
  return !!content.value?.text
})

const shapeStyle = computed(() => {
  return {
    width: '100%',
    height: '100%',
    position: 'relative'
  }
})

const textStyle = computed(() => {
  return {
    position: 'absolute',
    top: '50%',
    left: '50%',
    transform: 'translate(-50%, -50%)',
    width: '80%',
    height: '80%'
  }
})
</script>

<style scoped>
.shape-renderer {
  display: flex;
  align-items: center;
  justify-content: center;
}

.shape-renderer svg {
  display: block;
}

.shape-text {
  pointer-events: none;
}
</style>
