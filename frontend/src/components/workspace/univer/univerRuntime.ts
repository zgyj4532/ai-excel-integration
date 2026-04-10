export async function loadUniverRuntime() {
  const [corePreset, zhCNLocale, core, adapter, coreFacade] = await Promise.all([
    import('@univerjs/preset-sheets-core'),
    import('@univerjs/preset-sheets-core/locales/zh-CN'),
    import('@univerjs/core'),
    import('@univerjs/ui-adapter-vue3')
    , import('@univerjs/core/facade')
  ])

  await Promise.all([
    import('@univerjs/preset-sheets-core/lib/index.css'),
  ])

  return {
    Univer: core.Univer,
    FUniver: coreFacade.FUniver,
    CalculationMode: corePreset.CalculationMode,
    UniverSheetsCorePreset: corePreset.UniverSheetsCorePreset,
    UniverPresetSheetsCoreZhCN: zhCNLocale.default,
    LocaleType: core.LocaleType,
    mergeLocales: core.mergeLocales,
    greenTheme: (await import('@univerjs/themes')).greenTheme,
    UniverVue3AdapterPlugin: adapter.UniverVue3AdapterPlugin
  }
}