<template>
  <component :is="TemplatesPageAsync" @template-response="$emit('template-response', $event)" />
</template>

<script setup lang="ts">
import { defineAsyncComponent } from 'vue'

const TemplatesPageAsync = defineAsyncComponent(() => import('../../../pages/Templates.vue'))

defineEmits<{
  (e: 'template-response', payload: string): void
}>()
</script>
