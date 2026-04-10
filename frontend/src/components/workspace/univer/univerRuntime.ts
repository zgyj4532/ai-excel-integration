export async function loadUniverRuntime() {
  const [corePresetMod, zhCNLocaleMod, coreMod, adapterMod, coreFacadeMod] = await Promise.all([
    import('@univerjs/preset-sheets-core'),
    import('@univerjs/preset-sheets-core/locales/zh-CN'),
    import('@univerjs/core'),
    import('@univerjs/ui-adapter-vue3'),
    import('@univerjs/core/facade')
  ])

  await Promise.all([
    import('@univerjs/preset-sheets-core/lib/index.css'),
  ])

  const corePreset = corePresetMod?.default ?? corePresetMod
  const zhCNLocale = zhCNLocaleMod?.default ?? zhCNLocaleMod
  const core = coreMod?.default ?? coreMod
  const adapter = adapterMod?.default ?? adapterMod
  const coreFacade = coreFacadeMod?.default ?? coreFacadeMod

  const themes = await import('@univerjs/themes')
  const greenTheme = themes?.greenTheme ?? themes?.default?.greenTheme

  return {
    Univer: core.Univer,
    FUniver: coreFacade.FUniver,
    CalculationMode: corePreset.CalculationMode ?? corePreset?.default?.CalculationMode,
    UniverSheetsCorePreset: corePreset.UniverSheetsCorePreset ?? corePreset?.default?.UniverSheetsCorePreset,
    UniverPresetSheetsCoreZhCN: zhCNLocale,
    LocaleType: core.LocaleType,
    mergeLocales: core.mergeLocales,
    greenTheme,
    // expose commonly used plugins from preset to allow explicit registration
    UniverSheetsPlugin: corePreset.UniverSheetsPlugin ?? corePreset?.default?.UniverSheetsPlugin,
    UniverSheetsUIPlugin: corePreset.UniverSheetsUIPlugin ?? corePreset?.default?.UniverSheetsUIPlugin,
    UniverFormulaEnginePlugin: corePreset.UniverFormulaEnginePlugin ?? corePreset?.default?.UniverFormulaEnginePlugin,
    UniverRenderEnginePlugin: corePreset.UniverRenderEnginePlugin ?? corePreset?.default?.UniverRenderEnginePlugin,
    UniverUIPlugin: corePreset.UniverUIPlugin ?? corePreset?.default?.UniverUIPlugin,
    UniverSheetsFormulaPlugin: corePreset.UniverSheetsFormulaPlugin ?? corePreset?.default?.UniverSheetsFormulaPlugin,
    UniverSheetsFormulaUIPlugin: corePreset.UniverSheetsFormulaUIPlugin ?? corePreset?.default?.UniverSheetsFormulaUIPlugin,
    UniverSheetsNumfmtPlugin: corePreset.UniverSheetsNumfmtPlugin ?? corePreset?.default?.UniverSheetsNumfmtPlugin,
    UniverSheetsNumfmtUIPlugin: corePreset.UniverSheetsNumfmtUIPlugin ?? corePreset?.default?.UniverSheetsNumfmtUIPlugin,
    // docs and rpc plugins
    UniverDocsPlugin: corePreset.UniverDocsPlugin ?? corePreset?.default?.UniverDocsPlugin,
    UniverDocsUIPlugin: corePreset.UniverDocsUIPlugin ?? corePreset?.default?.UniverDocsUIPlugin,
    UniverNetworkPlugin: corePreset.UniverNetworkPlugin ?? corePreset?.default?.UniverNetworkPlugin,
    UniverRPCMainThreadPlugin: corePreset.UniverRPCMainThreadPlugin ?? corePreset?.default?.UniverRPCMainThreadPlugin,
    UniverRPCWorkerThreadPlugin: corePreset.UniverRPCWorkerThreadPlugin ?? corePreset?.default?.UniverRPCWorkerThreadPlugin,
    UniverVue3AdapterPlugin:
      adapter.UniverVue3AdapterPlugin ?? adapter?.default?.UniverVue3AdapterPlugin ?? adapter
  }
}