<template>
  <section class="preview workspace-tabs card">
    <div class="tabs-bar">
      <div class="tab-buttons">
        <button :class="['tab-btn', activeTab === 'overview' ? 'active' : '']" @click="$emit('update:activeTab', 'overview')">
          {{ $t('dataPreview') }}
        </button>
        <button :class="['tab-btn', activeTab === 'analysis' ? 'active' : '']" @click="$emit('update:activeTab', 'analysis')">
          {{ $t('analysis') }}
        </button>
        <button :class="['tab-btn', activeTab === 'templates' ? 'active' : '']" @click="$emit('update:activeTab', 'templates')">
          {{ $t('templates') }}
        </button>
      </div>

      <div class="action-buttons">
        <button
          class="undo-btn"
          :disabled="!hasFile || undoing"
          @click="$emit('undo')"
          :title="$t('undoFile')"
          :aria-label="$t('undoFile')"
        >
          <svg class="undo-icon" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
            <path d="M9 14L4 9l5-5" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
            <path d="M20 20v-7a4 4 0 0 0-4-4H4" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
          <span>{{ undoing ? $t('undoing') : $t('undoFile') }}</span>
        </button>

        <button
          class="save-btn"
          :disabled="!hasFile || saving"
          @click="$emit('save')"
          :title="$t('saveFile')"
          :aria-label="$t('saveFile')"
        >
          <svg class="save-icon" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
            <path d="M6 3h11l3 3v13a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z" fill="none" stroke="currentColor" stroke-width="1.6"/>
            <path d="M14 3v5a1 1 0 0 1-1 1H7a1 1 0 0 1-1-1V3" fill="none" stroke="currentColor" stroke-width="1.6"/>
            <path d="M15 17H9m0-3h6" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
          </svg>
          <span>{{ saving ? $t('saving') : $t('saveFile') }}</span>
        </button>
      </div>
    </div>

    <div v-if="saveMessage" class="save-message">{{ saveMessage }}</div>
    <div v-if="undoMessage" class="undo-message">{{ undoMessage }}</div>

    <div class="tabs-body">
      <div v-if="activeTab === 'overview'" class="tab-panel">
        <slot name="overview" />
      </div>
      <div v-else-if="activeTab === 'analysis'" class="tab-panel">
        <slot name="analysis" />
      </div>
      <div v-else class="tab-panel">
        <slot name="templates" />
      </div>
    </div>
  </section>
</template>

<script setup lang="ts">
defineProps<{
  activeTab: string
  hasFile: boolean
  saving: boolean
  saveMessage: string
  undoing: boolean
  undoMessage: string
}>()

defineEmits<{
  (e: 'update:activeTab', value: string): void
  (e: 'save'): void
  (e: 'undo'): void
}>()
</script>

<style scoped>
.workspace-tabs {
  display: flex;
  flex-direction: column;
  gap: 12px;
  min-width: 0;
  min-height: 0;
}

.tabs-bar {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px 16px 0;
  justify-content: space-between;
  flex-wrap: wrap;
}

.tab-buttons {
  display: flex;
  gap: 8px;
}

.tab-btn {
  background: transparent;
  border: 1px solid transparent;
  padding: 6px 12px;
  border-radius: 6px;
  cursor: pointer;
  color: #cfcfd6;
  transition: background-color 0.18s ease, color 0.18s ease, border-color 0.18s ease, transform 0.18s ease;
}

.tab-btn.active {
  background: var(--accent);
  border-color: transparent;
  color: var(--text-on-accent);
}

.tab-btn:hover {
  filter: brightness(1.05);
  transform: translateY(-1px);
}

.action-buttons {
  margin-left: auto;
  display: flex;
  gap: 8px;
}

.undo-btn {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 8px 14px;
  border-radius: 10px;
  border: 1px solid rgba(100, 149, 237, 0.32);
  background: linear-gradient(135deg, rgba(100, 149, 237, 0.12), rgba(70, 130, 180, 0.12));
  color: var(--text-primary);
  cursor: pointer;
  box-shadow: 0 12px 28px rgba(0, 0, 0, 0.32);
  transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease, background 0.2s ease, opacity 0.2s ease;
}

.undo-btn:hover:not(:disabled) {
  transform: translateY(-1px);
  box-shadow: 0 16px 40px rgba(0, 0, 0, 0.4);
  border-color: rgba(100, 149, 237, 0.52);
  background: linear-gradient(135deg, rgba(100, 149, 237, 0.16), rgba(70, 130, 180, 0.18));
}

.undo-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.undo-icon {
  width: 18px;
  height: 18px;
}

.save-btn {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 8px 14px;
  border-radius: 10px;
  border: 1px solid rgba(16, 185, 129, 0.32);
  background: linear-gradient(135deg, rgba(25, 179, 148, 0.12), rgba(16, 185, 129, 0.12));
  color: var(--text-primary);
  cursor: pointer;
  box-shadow: 0 12px 28px rgba(0, 0, 0, 0.32);
  transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease, background 0.2s ease, opacity 0.2s ease;
}

.save-btn:hover:not(:disabled) {
  transform: translateY(-1px);
  box-shadow: 0 16px 40px rgba(0, 0, 0, 0.4);
  border-color: rgba(16, 185, 129, 0.52);
  background: linear-gradient(135deg, rgba(16, 185, 129, 0.16), rgba(25, 179, 148, 0.18));
}

.save-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.save-icon {
  width: 18px;
  height: 18px;
}

.tabs-body {
  padding: 0 16px 16px;
  min-height: 0;
}

.save-message {
  margin: 0 16px;
  color: #d9ead3;
  font-size: 13px;
}

.undo-message {
  margin: 0 16px;
  color: #b0c4de;
  font-size: 13px;
}

.tab-panel {
  min-height: 0;
}

@media (max-width: 800px) {
  .tabs-bar {
    align-items: flex-start;
    flex-direction: column;
  }

  .save-btn {
    margin-left: 0;
  }
}
</style>
