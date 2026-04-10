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

    <div v-if="saveMessage" class="save-message">{{ saveMessage }}</div>

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
}>()

defineEmits<{
  (e: 'update:activeTab', value: string): void
  (e: 'save'): void
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
  background: #189079;
  border-color: transparent;
  color: #fff;
}

.tab-btn:hover {
  filter: brightness(1.05);
  transform: translateY(-1px);
}

.save-btn {
  margin-left: auto;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 8px 14px;
  border-radius: 10px;
  border: 1px solid rgba(197, 160, 89, 0.32);
  background: linear-gradient(135deg, rgba(25, 179, 148, 0.12), rgba(197, 160, 89, 0.12));
  color: #e0e0e0;
  cursor: pointer;
  box-shadow: 0 12px 28px rgba(0, 0, 0, 0.32);
  transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease, background 0.2s ease, opacity 0.2s ease;
}

.save-btn:hover:not(:disabled) {
  transform: translateY(-1px);
  box-shadow: 0 16px 40px rgba(0, 0, 0, 0.4);
  border-color: rgba(197, 160, 89, 0.52);
  background: linear-gradient(135deg, rgba(197, 160, 89, 0.16), rgba(25, 179, 148, 0.18));
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
