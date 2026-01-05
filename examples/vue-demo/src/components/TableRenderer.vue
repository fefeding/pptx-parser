<template>
  <div class="table-renderer" :style="tableStyle">
    <table v-if="rows.length > 0">
      <tbody>
        <tr v-for="(row, rowIndex) in rows" :key="rowIndex">
          <td v-for="(cell, cellIndex) in row.cells" :key="cellIndex" :style="getCellStyle(cell)">
            <TextRenderer :element="{ id: `${rowIndex}-${cellIndex}`, type: 'text', x: 0, y: 0, width: cell.width, height: cell.height, zIndex: 1, content: cell.content }" />
          </td>
        </tr>
      </tbody>
    </table>
    <div v-else class="placeholder">表格</div>
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

const rows = computed(() => {
  if (props.element.type !== 'table') return []
  return (props.element.content as any)?.rows || []
})

const tableStyle = computed(() => {
  const style: any = {
    width: '100%',
    height: '100%',
    borderCollapse: 'collapse'
  }

  const element = props.element as any
  if (element.style?.backgroundColor) {
    style.backgroundColor = element.style.backgroundColor
  }

  return style
})

function getCellStyle(cell: any) {
  const style: any = {
    border: '1px solid #ccc',
    padding: '8px',
    verticalAlign: 'middle'
  }

  if (cell.width) {
    style.width = `${cell.width}px`
  }

  if (cell.height) {
    style.height = `${cell.height}px`
  }

  if (cell.style?.backgroundColor) {
    style.backgroundColor = cell.style.backgroundColor
  }

  if (cell.style?.textAlign) {
    style.textAlign = cell.style.textAlign
  }

  return style
}
</script>

<style scoped>
.table-renderer {
  overflow: auto;
}

.table-renderer table {
  width: 100%;
  height: 100%;
  border: 1px solid #ccc;
}

.table-renderer td {
  overflow: hidden;
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
