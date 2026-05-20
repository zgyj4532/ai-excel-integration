<template>
  <div class="report-actions">
    <div class="report-buttons">
      <button
        @click="$emit('generate-report')"
        class="generate-report-btn"
        :disabled="reportGenerating || downloadInProgress"
      >{{ reportGenerating ? '信息生成中...' : $t('generateReportBtn') }}</button>
      <button @click="$emit('load-api-example')" class="generate-report-btn" style="margin-left:8px">{{ $t('loadApiExample') }}</button>
      <button
        v-if="reportGenerated"
        @click="$emit('download-report')"
        class="generate-report-btn download-report-btn"
        :disabled="reportGenerating || downloadInProgress"
      >{{ downloadInProgress ? '下载中，请稍后' : '下载报告' }}</button>
    </div>

    <div class="report-options">
      <span class="options-label">{{ $t('optionalAnalyses') }}</span>
      <label><input type="checkbox" v-model="financialRatiosModel" />{{ $t('includeFinancialRatios') }}</label>
      <label><input type="checkbox" v-model="profitabilityModel" />{{ $t('includeProfitability') }}</label>
      <label><input type="checkbox" v-model="cashFlowModel" />{{ $t('includeCashFlow') }}</label>
      <label><input type="checkbox" v-model="budgetActualModel" />{{ $t('includeBudgetActual') }}</label>
      <label><input type="checkbox" v-model="rfmModel" />{{ $t('includeRfm') }}</label>
      <label><input type="checkbox" v-model="clvModel" />{{ $t('includeClv') }}</label>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue'

const props = defineProps<{
  reportGenerated: boolean
  reportGenerating: boolean
  downloadInProgress: boolean
  includeFinancialRatios: boolean
  includeProfitability: boolean
  includeCashFlow: boolean
  includeBudgetActual: boolean
  includeRfm: boolean
  includeClv: boolean
}>()

const emit = defineEmits<{
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

const financialRatiosModel = computed({
  get: () => props.includeFinancialRatios,
  set: value => emit('update:includeFinancialRatios', value)
})

const profitabilityModel = computed({
  get: () => props.includeProfitability,
  set: value => emit('update:includeProfitability', value)
})

const cashFlowModel = computed({
  get: () => props.includeCashFlow,
  set: value => emit('update:includeCashFlow', value)
})

const budgetActualModel = computed({
  get: () => props.includeBudgetActual,
  set: value => emit('update:includeBudgetActual', value)
})

const rfmModel = computed({
  get: () => props.includeRfm,
  set: value => emit('update:includeRfm', value)
})

const clvModel = computed({
  get: () => props.includeClv,
  set: value => emit('update:includeClv', value)
})
</script>

<style scoped>
.report-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
  margin-top: 12px;
}

.report-buttons {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
}

.generate-report-btn {
  background: #10b981;
  color: var(--text-on-accent);
  border: 0;
  padding: 10px 16px;
  border-radius: 10px;
  cursor: pointer;
  margin-top: 0;
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

.download-report-btn {
  margin-left: auto;
  white-space: nowrap;
}

.report-options {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  align-items: center;
}

.options-label {
  color: var(--text-dim);
  font-size: 12px;
}

.report-options label {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: var(--text-primary);
  font-size: 13px;
}

.report-options input[type='checkbox'] {
  accent-color: #19b394;
}
</style>