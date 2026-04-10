<template>
  <div class="workspace-sidebar card">
    <div v-if="showUploader" class="uploader-wrap">
      <FileUploader :snapshotOps="snapshotOps" @fileLoaded="$emit('file-loaded', $event)" />
    </div>

    <div class="chat-box">
      <div v-if="aiActive" class="chat-stream">
        <ChatBubbleList
          :messages="aiMessages"
          :executingIds="executingIds"
          :appliedIds="appliedIds"
          :snapshot-ops="snapshotOps"
          @apply-token="$emit('apply-token', $event)"
          @skip-token="$emit('skip-token', $event)"
        />
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import FileUploader from '../FileUploader.vue'
import ChatBubbleList from '../ChatBubbleList.vue'

defineProps<{
  showUploader: boolean
  aiActive: boolean
  aiMessages: Array<{ id: number; role: string; text: string; placeholder?: boolean }>
  executingIds: string[]
  appliedIds: string[]
  snapshotOps: any
}>()

defineEmits<{
  (e: 'file-loaded', payload: { name: string; data: string[][]; file?: File }): void
  (e: 'apply-token', token: string, msgId?: number, idx?: number): void
  (e: 'skip-token', tokenKey: string, msgId?: number, idx?: number, note?: string): void
}>()
</script>

<style scoped>
.workspace-sidebar {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.chat-box {
  display: flex;
  flex-direction: column;
  gap: 8px;
  min-height: 0;
}

.chat-stream {
  min-height: 0;
}
</style>
