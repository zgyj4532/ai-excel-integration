<template>
  <div class="workspace-header">
    <div class="muted">{{ $t('currentFile') }} {{ fileName || $t('noFile') }}</div>
    <div class="actions">
      <button class="new-chat-btn" @click="$emit('new-chat')" :title="t('newChat')" :aria-label="t('newChat')">
        <span class="new-chat-circle" aria-hidden="true">
          <svg class="new-chat-plus" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" role="img" aria-hidden="true">
            <line x1="12" y1="5" x2="12" y2="19" stroke="#179078" stroke-width="2" stroke-linecap="round" />
            <line x1="5" y1="12" x2="19" y2="12" stroke="#179078" stroke-width="2" stroke-linecap="round" />
          </svg>
        </span>
        <span class="new-chat-label">{{ t('newChat') }}</span>
      </button>
    </div>
    <div v-if="lastError" class="error-banner">
      <div class="error-text">{{ lastError }}</div>
      <button class="btn retry-btn" @click="$emit('retry')" v-if="failedCommand">{{ t('retry') || '重试' }}</button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { useI18n } from 'vue-i18n'

const { t } = useI18n()

defineProps<{
  fileName: string
  lastError: string | null
  failedCommand: string | null
}>()

defineEmits<{
  (e: 'new-chat'): void
  (e: 'retry'): void
}>()
</script>

<style scoped>
.workspace-header {
  width: 100%;
  flex: 1 1 auto;
  align-self: stretch;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.workspace-header > .muted {
  width: 100%;
  text-align: left;
}

.actions {
  display: flex;
  gap: 8px;
  align-items: center;
}

.workspace-header > .actions {
  margin-left: auto;
}

.error-banner {
  display: flex;
  gap: 8px;
  align-items: center;
  width: 100%;
}

.error-text {
  color: #ffcccc;
  background: rgba(255, 0, 0, 0.06);
  padding: 8px;
  border-radius: 6px;
  flex: 1;
}
</style>
