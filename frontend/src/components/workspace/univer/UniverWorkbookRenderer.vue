<template>
  <section class="univer-table">
    <div :id="containerId" class="univer-container"></div>
  </section>
</template>

<script setup lang="ts">
import { onBeforeUnmount, onMounted } from 'vue'
import { loadUniverRuntime } from './univerRuntime'

const emit = defineEmits<{ (e: 'ready', payload: { univerAPI: any }): void }>()

const containerId = 'univer-table-container'
let disposeUniver: (() => void) | null = null

onMounted(async () => {
  try {
    const {
      FUniver,
      Univer,
      CalculationMode,
      UniverSheetsCorePreset,
      UniverPresetSheetsCoreZhCN,
      LocaleType,
      mergeLocales,
      greenTheme,
      UniverVue3AdapterPlugin,
    } = await loadUniverRuntime()

    const univer = new Univer({
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
    } as any)

    univer.registerPlugin(UniverVue3AdapterPlugin)

    const univerAPI = FUniver.newAPI(univer)
    univerAPI.createWorkbook({})

    emit('ready', { univerAPI })
    disposeUniver = () => univer.dispose()
  } catch (error) {
    console.error('Failed to initialize Univer:', error)
  }
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