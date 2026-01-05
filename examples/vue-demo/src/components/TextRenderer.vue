<template>
  <div class="text-renderer" :style="textStyle">
    <div v-for="(paragraph, pIndex) in paragraphs" :key="pIndex" class="paragraph" :style="paragraphStyle">
      <span
        v-for="(run, rIndex) in paragraph.runs"
        :key="`${pIndex}-${rIndex}`"
        class="text-run"
        :style="getRunStyle(run)"
      >
        {{ run.text }}
      </span>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import type { PptNode } from 'pptx-parser'

interface Props {
  element: PptNode
}

const props = defineProps<Props>()

const paragraphs = computed(() => {
  if (props.element.type !== 'text') return []
  return props.element.content?.paragraphs || []
})

const textStyle = computed(() => {
  const style: any = {
    width: '100%',
    height: '100%',
    overflow: 'hidden'
  }

  const element = props.element as any
  if (element.style?.textAlign) {
    style.textAlign = element.style.textAlign
  }

  if (element.style?.verticalAlign) {
    style.display = 'flex'
    style.flexDirection = 'column'
    style.justifyContent = element.style.verticalAlign === 'middle' ? 'center' :
                         element.style.verticalAlign === 'bottom' ? 'flex-end' : 'flex-start'
  }

  return style
})

const paragraphStyle = computed(() => {
  const style: any = {}

  const element = props.element as any
  if (element.style?.lineHeight) {
    style.lineHeight = element.style.lineHeight
  }

  return style
})

function getRunStyle(run: any) {
  const style: any = {}

  if (run.fontFamily) {
    style.fontFamily = run.fontFamily
  }

  if (run.fontSize) {
    style.fontSize = `${run.fontSize}px`
  }

  if (run.fontColor) {
    style.color = run.fontColor
  }

  if (run.bold) {
    style.fontWeight = 'bold'
  }

  if (run.italic) {
    style.fontStyle = 'italic'
  }

  if (run.underline) {
    style.textDecoration = 'underline'
  }

  if (run.strike) {
    style.textDecoration = (style.textDecoration || '') + ' line-through'
  }

  return style
}
</script>

<style scoped>
.text-renderer {
  word-wrap: break-word;
  white-space: pre-wrap;
}

.paragraph {
  margin: 0;
}

.text-run {
  display: inline;
}
</style>
