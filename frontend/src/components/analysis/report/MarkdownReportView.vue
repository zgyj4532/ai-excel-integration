<template>
  <div class="markdown-report-view" v-html="html"></div>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import { renderMarkdown } from './reportMarkdown'

const props = defineProps<{
  markdown: string
}>()

const html = computed(() => renderMarkdown(props.markdown))
</script>

<style scoped>
.markdown-report-view {
  color: var(--text-primary);
  line-height: 1.7;
  word-break: break-word;
}

.markdown-report-view :deep(h1),
.markdown-report-view :deep(h2),
.markdown-report-view :deep(h3),
.markdown-report-view :deep(h4),
.markdown-report-view :deep(h5),
.markdown-report-view :deep(h6) {
  margin: 0.8em 0 0.35em;
  line-height: 1.25;
}

.markdown-report-view :deep(p) {
  margin: 0.45em 0;
}

.markdown-report-view :deep(ul),
.markdown-report-view :deep(ol) {
  margin: 0.5em 0 0.75em 1.35em;
  padding: 0;
}

.markdown-report-view :deep(code) {
  padding: 0.15em 0.35em;
  border-radius: 6px;
  background: rgba(255, 255, 255, 0.08);
  font-family: 'IBM Plex Mono', monospace;
  font-size: 0.94em;
}

.markdown-report-view :deep(a) {
  color: #8bd5ca;
  text-decoration: underline;
}
</style>