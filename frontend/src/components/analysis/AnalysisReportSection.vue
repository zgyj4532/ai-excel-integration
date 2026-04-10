<template>
  <div class="analysis-report-section">
    <div class="card auto-report-card" :class="{ generated: reportGenerated }">
      <h4>{{ $t('autoReportTitle') }}</h4>
      <p class="muted">{{ $t('autoReportDesc') }}</p>

      <AnalysisReportControls
        :report-generated="reportGenerated"
        :report-generating="reportGenerating"
        :download-in-progress="downloadInProgress"
        :include-financial-ratios="includeFinancialRatios"
        :include-profitability="includeProfitability"
        :include-cash-flow="includeCashFlow"
        :include-budget-actual="includeBudgetActual"
        :include-rfm="includeRfm"
        :include-clv="includeClv"
        @generate-report="$emit('generate-report')"
        @load-api-example="$emit('load-api-example')"
        @download-report="$emit('download-report')"
        @update:includeFinancialRatios="$emit('update:includeFinancialRatios', $event)"
        @update:includeProfitability="$emit('update:includeProfitability', $event)"
        @update:includeCashFlow="$emit('update:includeCashFlow', $event)"
        @update:includeBudgetActual="$emit('update:includeBudgetActual', $event)"
        @update:includeRfm="$emit('update:includeRfm', $event)"
        @update:includeClv="$emit('update:includeClv', $event)"
      />

      <div
        class="auto-report-body"
        :style="{
          height: reportGenerated ? 'auto' : '160px',
          maxHeight: reportGenerated ? '720px' : '160px',
          overflowY: reportGenerated ? 'auto' : 'hidden'
        }"
      >
        <div v-if="!reportGenerated" class="report-placeholder">
          {{ $t('reportPlaceholder') }}
        </div>
        <div v-else class="report-content">
          <MarkdownReportView :markdown="reportMarkdown" />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import AnalysisReportControls from './report/AnalysisReportControls.vue'
import MarkdownReportView from './report/MarkdownReportView.vue'

defineProps<{
  reportGenerated: boolean
  reportGenerating: boolean
  downloadInProgress: boolean
  reportMarkdown: string
  includeFinancialRatios: boolean
  includeProfitability: boolean
  includeCashFlow: boolean
  includeBudgetActual: boolean
  includeRfm: boolean
  includeClv: boolean
}>()

defineEmits<{
  (e: 'generate-report'): void
  (e: 'load-api-example'): void
  (e: 'download-report'): void
  (e: 'update:includeFinancialRatios', value: boolean): void
  (e: 'update:includeProfitability', value: boolean): void
  (e: 'update:includeCashFlow', value: boolean): void
  (e: 'update:includeBudgetActual', value: boolean): void
  (e: 'update:includeRfm', value: boolean): void
  (e: 'update:includeClv', value: boolean): void
}>()
</script>

<style scoped>
.analysis-report-section {
  min-width: 0;
  animation: slideUpReveal 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards;
}

.auto-report-card {
  min-height: 160px;
  padding: 8px;
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid rgba(255, 255, 255, 0.08);
  border-radius: 12px;
  box-shadow: 0 12px 32px rgba(0, 0, 0, 0.35);
  animation: slideUpReveal 0.9s cubic-bezier(0.16, 1, 0.3, 1) forwards;
}

.auto-report-body {
  margin-top: 12px;
}

.report-placeholder {
  display: flex;
  align-items: center;
  justify-content: center;
  height: 100%;
  color: rgba(230, 238, 248, 0.7);
}

.report-content {
  min-height: 100%;
}

.analysis-report-section :deep(h4) {
  margin: 0;
  font-size: 16px;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #c5a059;
}

.analysis-report-section :deep(.muted) {
  margin-top: 4px;
  font-size: 12px;
  color: rgba(230, 238, 248, 0.72);
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
</style>
