<template>
  <component :is="UniverWorkbookRendererAsync" @ready="$emit('ready', $event)" />
</template>

<script setup lang="ts">
import { defineAsyncComponent } from 'vue'

const UniverWorkbookRendererAsync = defineAsyncComponent(() => import('./workspace/univer/UniverWorkbookRenderer.vue'))

defineEmits<{
  (e: 'ready', payload: { univerAPI: any }): void
}>()
</script>
