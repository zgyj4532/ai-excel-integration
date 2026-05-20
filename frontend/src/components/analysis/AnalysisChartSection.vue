<template>
  <div class="analysis-chart-section">
    <div class="card charts-row">
      <div class="card-header">
        <h4>{{ $t('chartSuggestionsTitle') }}</h4>
        <p class="muted">{{ $t('chartSuggestionsDesc') }}</p>
      </div>

      <div class="data-range-row">
        <label class="range-label">{{ $t('dataRangeLabel') }}</label>
        <input
          :value="dataRangeInput"
          @input="onRangeInput"
          class="range-input"
          :placeholder="$t('dataRangePlaceholder')"
        />
        <span class="range-hint">{{ $t('rangeHint') }}</span>
        <span v-if="rangeError" class="range-error">{{ rangeError }}</span>
      </div>

      <div class="chart-suggestions-grid horizontal">
        <div class="chart-suggestion-item">
          <h5>{{ $t('chart_line') }}</h5>
          <div class="chart-placeholder">
            <div class="range-preview">{{ $t('dataRangeLabel') }}：{{ dataRangeInput || $t('rangeNotDetected') }}</div>
            <div class="target-row">
              <label class="range-label">{{ $t('targetColumnLabel') }}</label>
              <select v-model="lineModel">
                <option v-for="(h, idx) in sheetHeaders" :key="idx" :value="h">{{ h || $t('col_default', { n: idx + 1 }) }}</option>
              </select>
            </div>
            <div class="action-row">
              <button @click="$emit('create-chart', 'line', lineModel)" class="generate-report-btn" :disabled="!savedFile">{{ $t('createChartBtn') }}</button>
            </div>
            <div class="preview-area">
              <canvas ref="lineCanvasRef" class="chart-canvas"></canvas>
              <div v-if="chartInstructionsTextLine" class="chart-instructions">{{ chartInstructionsTextLine }}</div>
            </div>
          </div>
        </div>

        <div class="chart-suggestion-item">
          <h5>{{ $t('chart_pie') }}</h5>
          <div class="chart-placeholder">
            <div class="range-preview">{{ $t('dataRangeLabel') }}：{{ dataRangeInput || $t('rangeNotDetected') }}</div>
            <div class="target-row">
              <label class="range-label">{{ $t('targetColumnLabel') }}</label>
              <select v-model="pieModel">
                <option v-for="(h, idx) in sheetHeaders" :key="idx" :value="h">{{ h || $t('col_default', { n: idx + 1 }) }}</option>
              </select>
            </div>
            <div class="action-row">
              <button @click="$emit('create-chart', 'pie', pieModel)" class="generate-report-btn" :disabled="!savedFile">{{ $t('createChartBtn') }}</button>
            </div>
            <div class="preview-area">
              <canvas ref="pieCanvasRef" class="chart-canvas"></canvas>
              <div v-if="chartInstructionsTextPie" class="chart-instructions">{{ chartInstructionsTextPie }}</div>
            </div>
          </div>
        </div>

        <div class="chart-suggestion-item">
          <h5>{{ $t('chart_top') }}</h5>
          <div class="chart-placeholder">
            <div class="range-preview">{{ $t('dataRangeLabel') }}：{{ dataRangeInput || $t('rangeNotDetected') }}</div>
            <div class="target-row">
              <label class="range-label">{{ $t('targetColumnLabel') }}</label>
              <select v-model="barModel">
                <option v-for="(h, idx) in sheetHeaders" :key="idx" :value="h">{{ h || $t('col_default', { n: idx + 1 }) }}</option>
              </select>
            </div>
            <div class="action-row">
              <button @click="$emit('create-chart', 'bar', barModel)" class="generate-report-btn" :disabled="!savedFile">{{ $t('createChartBtn') }}</button>
            </div>
            <div class="preview-area">
              <canvas ref="barCanvasRef" class="chart-canvas"></canvas>
              <div v-if="chartInstructionsTextBar" class="chart-instructions">{{ chartInstructionsTextBar }}</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from 'vue'

const props = defineProps<{
  dataRangeInput: string
  rangeError: string
  savedFile: boolean
  sheetHeaders: string[]
  selectedColumnLine: string
  selectedColumnPie: string
  selectedColumnBar: string
  chartInstructionsTextLine: string
  chartInstructionsTextPie: string
  chartInstructionsTextBar: string
}>()

const emit = defineEmits<{
  (e: 'input'): void
  (e: 'create-chart', chartType: string, targetColumn: string): void
  (e: 'update:selectedColumnLine', value: string): void
  (e: 'update:selectedColumnPie', value: string): void
  (e: 'update:selectedColumnBar', value: string): void
}>()

const lineModel = computed({
  get: () => props.selectedColumnLine,
  set: value => emit('update:selectedColumnLine', value)
})

const pieModel = computed({
  get: () => props.selectedColumnPie,
  set: value => emit('update:selectedColumnPie', value)
})

const barModel = computed({
  get: () => props.selectedColumnBar,
  set: value => emit('update:selectedColumnBar', value)
})

const lineCanvasRef = ref<HTMLCanvasElement | null>(null)
const pieCanvasRef = ref<HTMLCanvasElement | null>(null)
const barCanvasRef = ref<HTMLCanvasElement | null>(null)

function onRangeInput() {
  emit('input')
}

defineExpose({
  lineCanvasRef,
  pieCanvasRef,
  barCanvasRef
})
</script>

<style scoped>
.analysis-chart-section {
  min-width: 0;
  animation: slideUpReveal 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards;
}

.card-header {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.data-range-row {
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-bottom: 16px;
}

.range-label {
  font-size: 12px;
  color: var(--text-dim);
}

.range-input {
  max-width: 320px;
  background: var(--input-bg);
  border: 1px solid rgba(255, 255, 255, 0.08);
  color: var(--text-primary);
  padding: 6px 10px;
  border-radius: 8px;
  width: 140px;
  font-family: 'IBM Plex Mono', 'Space Grotesk', monospace;
}

.range-hint,
.range-error,
.range-preview {
  font-size: 12px;
}

.range-hint {
  color: rgba(200, 210, 220, 0.7);
}

.range-error {
  color: #fca5a5;
}

.chart-suggestions-grid.horizontal {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 16px;
}

.chart-suggestion-item {
  border: 1px solid rgba(255, 255, 255, 0.06);
  border-radius: 12px;
  padding: 12px;
  background: rgba(255, 255, 255, 0.02);
  min-width: 0;
  box-shadow: 0 12px 32px rgba(0, 0, 0, 0.35);
  animation: slideUpReveal 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards;
}

.generate-report-btn {
  background: #10b981;
  color: var(--text-on-accent);
  border: 0;
  padding: 10px 16px;
  border-radius: 10px;
  cursor: pointer;
  margin-top: 12px;
  font-weight: 700;
  letter-spacing: 0.02em;
  box-shadow: 0 10px 28px rgba(16, 185, 129, 0.26);
  transition: transform 0.2s ease, box-shadow 0.2s ease, opacity 0.2s ease;
  font-family: 'Space Grotesk', 'IBM Plex Mono', system-ui, sans-serif;
}

.generate-report-btn:hover:not(:disabled) {
  transform: translateY(-2px);
  box-shadow: 0 14px 34px rgba(0, 0, 0, 0.35);
}

.generate-report-btn:disabled {
  opacity: 0.55;
  cursor: not-allowed;
}

.chart-placeholder {
  display: flex;
  flex-direction: column;
  gap: 8px;
  align-items: stretch;
}

.target-row,
.action-row,
.preview-area {
  display: flex;
  gap: 8px;
  align-items: center;
  flex-wrap: wrap;
}

.preview-area {
  margin-top: 6px;
  color: var(--text-primary);
  font-size: 13px;
  flex-direction: column;
  align-items: stretch;
}

.chart-canvas {
  width: 100%;
  height: 220px;
}

.chart-instructions {
  color: var(--text-primary);
}

.analysis-chart-section :deep(h4) {
  margin: 0;
  font-size: 16px;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #10b981;
}

.analysis-chart-section :deep(.muted) {
  margin-top: 4px;
  font-size: 12px;
  color: var(--text-dim);
}

@keyframes slideUpReveal {
  from {
    opacity: 0;
    transform: translateY(18px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

@media (max-width: 960px) {
  .chart-suggestions-grid.horizontal {
    grid-template-columns: 1fr;
  }
}
</style>
