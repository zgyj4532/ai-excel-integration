<template>
  <div class="overview-tab">
    <h4 class="tab-heading">{{ t('dataPreview') }}</h4>
    <div v-if="hasFile" class="univer-wrapper">
      <component :is="UniverTableAsync" @ready="$emit('ready', $event)" />
    </div>
    <div v-else class="empty-upload">
      <div class="empty-upload-text">{{ t('workspaceNoUploadTip') }}</div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { defineAsyncComponent } from 'vue'
import { useI18n } from 'vue-i18n'

const { t } = useI18n()
const UniverTableAsync = defineAsyncComponent(() => import('../../UniverTable.vue'))

defineProps<{ hasFile: boolean }>()

defineEmits<{
  (e: 'ready', payload: { univerAPI: any }): void
}>()
</script>

<style scoped>
.overview-tab {
  min-width: 0;
}

.tab-heading {
  margin-top: 0;
}

.univer-wrapper {
  height: 610px;
  border: 1px solid #ddd;
  border-radius: 8px;
  overflow: hidden;
}

.empty-upload {
  height: 610px;
  border: 1px dashed rgba(16, 185, 129, 0.32);
  border-radius: 10px;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(255, 255, 255, 0.02);
  color: var(--text-primary);
  text-align: center;
}

.empty-upload-text {
  max-width: 420px;
  line-height: 1.6;
  font-size: 24px;
}
</style>
