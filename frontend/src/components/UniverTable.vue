<template>
  <section class="univer-table">
    <div :id="containerId" class="univer-container"></div>
  </section>
</template>

<script setup lang="ts">
import { onBeforeUnmount, onMounted } from 'vue'
import { UniverSheetsCorePreset, CalculationMode } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import { createUniver, LocaleType, mergeLocales, greenTheme } from '@univerjs/presets'
import { UniverVue3AdapterPlugin } from '@univerjs/ui-adapter-vue3'
import '@univerjs/preset-sheets-core/lib/index.css'
import '@univerjs/engine-formula/facade'
import '@univerjs/sheets-formula/facade'

const emit = defineEmits<{ (e: 'ready', payload: { univerAPI: any }): void }>()

const containerId = 'univer-table-container'
let disposeUniver: (() => void) | null = null

onMounted(() => {
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    darkMode: true,
    theme: greenTheme,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN),
    },
    presets: [
      UniverSheetsCorePreset({
        container: containerId,
        formula: {
          initialFormulaComputing: CalculationMode.FORCED,
        },
      }),
    ],
  })

  univer.registerPlugin(UniverVue3AdapterPlugin)

  univerAPI.createWorkbook({})

  emit('ready', { univerAPI })

  disposeUniver = () => univer.dispose()
})

onBeforeUnmount(() => {
  disposeUniver?.()
})
</script>

<style scoped>
.univer-table {
  height: 100%;
  margin: 0;
  background: #111827;
}

.univer-container {
  height: 100%;
  background: transparent;
}
</style>