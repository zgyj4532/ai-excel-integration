<template>
  <component :is="AnalysisPageAsync" :saved-file-id="savedFileId" :last-file="lastFile" />
</template>

<script setup lang="ts">
import { defineAsyncComponent } from 'vue'

const AnalysisPageAsync = defineAsyncComponent(() => import('../../../pages/Analysis.vue'))

defineProps<{
  savedFileId?: string | null
  lastFile?: File | null
}>()
</script>
