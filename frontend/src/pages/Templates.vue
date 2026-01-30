<template>
  <div class="templates">
    <div class="topbar">
      <h2>{{ $t('templatesTitle') }}</h2>
    </div>

    <div class="content">
      <!-- 第一行：模板管理 -->
      <div class="row row-first">
        <div class="card">
          <h4>{{ $t('template_management_title') }}</h4>
          <p class="muted">{{ $t('template_management_subtitle') }}</p>
          
          <div class="template-actions">
            <button @click="onApplyTemplate" class="apply-template-btn">{{ $t('applyTemplateBtn') }}</button>
            <button @click="onExportTemplates" class="export-btn">{{ $t('exportTemplates') }}</button>
            <button @click="onImportTemplates" class="import-btn">{{ $t('importTemplates') }}</button>
          </div>
        </div>
      </div>

      <!-- 标准模板库：来源于接口文档第7章高级AI操作 -->
      <div class="row row-library">
        <div class="card">
          <div class="library-header">
            <div>
              <h4>{{ $t('template_standard_library_title') }}</h4>
            </div>
            <div class="muted small">{{ $t('template_standard_library_count', { count: standardTemplates.length }) }}</div>
          </div>

          <div class="library-grid">
            <div v-for="tpl in standardTemplates" :key="tpl.id" class="library-card">
              <div class="library-card-header">
                <div class="library-title">{{ tpl.title }}</div>
                <span class="endpoint">{{ tpl.endpoint }}</span>
              </div>
              <p class="muted">{{ tpl.desc }}</p>
              <label class="muted" style="font-size:12px">Instructions</label>
              <textarea
                class="json-preview input-area"
                v-model="instructionInputs[tpl.id]"
                :placeholder="tpl.payload?.instructions || t('template_enter_instructions') || '请输入指令'"
                rows="3"
              ></textarea>
              <div class="template-actions compact">
                <button
                  @click="applyStandardTemplate(tpl)"
                  class="apply-btn"
                  :disabled="applyingId === tpl.id"
                >{{ applyingId === tpl.id ? '文本应用中...' : $t('applyThisTemplate') }}</button>
                <button @click="exportTemplate(tpl)" class="export-btn">{{ $t('exportTemplates') }}</button>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- 示例：Excel 数据分析模板 JSON -->
      <!-- <div class="row row-example">
        <div class="card">
          <div class="library-card-header">
            <div class="library-title">{{ $t('template_excel_analyze_title') }}</div>
            <span class="endpoint">{{ excelAnalyzeTemplate.endpoint }}</span>
          </div>
          <p class="muted">{{ $t('template_excel_analyze_desc') }}</p>
          <pre class="json-preview">{{ formatJson(templatePayload(excelAnalyzeTemplate)) }}</pre>
          <div class="template-actions compact">
            <button @click="applyStandardTemplate(excelAnalyzeTemplate)" class="apply-btn">{{ $t('applyThisTemplate') }}</button>
            <button @click="copyJson(templatePayload(excelAnalyzeTemplate))" class="export-btn">{{ $t('template_copy_json') }}</button>
          </div>
        </div>
      </div> -->
    </div>
  </div>
</template>

<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import { applyTemplate } from '@/services/api'
import { ref, onMounted, watch } from 'vue'

const { t } = useI18n()
const emit = defineEmits<{
  (e: 'template-response', payload: string): void
}>()

type StandardTemplate = {
  id: number
  title: string
  desc: string
  endpoint: string
  method?: string
  payload: Record<string, any>
}

const instructionInputs = ref<Record<number, string>>({})
const applyingId = ref<number | null>(null)

const standardTemplates = ref<StandardTemplate[]>([
  {
    id: 101,
    title: t('template_smart_cleaning_title'),
    desc: t('template_smart_cleaning_desc'),
    endpoint: '/api/ai/smart-data-cleaning',
    method: 'POST',
    payload: { instructions: '删除空行和空列，标准化日期格式，去除重复记录' }
  },
  {
    id: 102,
    title: t('template_smart_transform_title'),
    desc: t('template_smart_transform_desc'),
    endpoint: '/api/ai/smart-data-transformation',
    method: 'POST',
    payload: { instructions: '创建透视表，按部门汇总销售额，添加同比增长率列' }
  },
  {
    id: 103,
    title: t('template_smart_analysis_title'),
    desc: t('template_smart_analysis_desc'),
    endpoint: '/api/ai/smart-data-analysis',
    method: 'POST',
    payload: { instructions: '分析销售趋势，识别季节性模式，预测下季度销售额' }
  },
  {
    id: 104,
    title: t('template_smart_chart_title'),
    desc: t('template_smart_chart_desc'),
    endpoint: '/api/ai/smart-chart-creation',
    method: 'POST',
    payload: { instructions: '创建销售趋势折线图，包含预算线对比，添加数据标签' }
  },
  {
    id: 105,
    title: t('template_smart_validation_title'),
    desc: t('template_smart_validation_desc'),
    endpoint: '/api/ai/smart-data-validation',
    method: 'POST',
    payload: { instructions: '验证数据完整性，检查数值范围，标记异常值' }
  }
])

const excelAnalyzeTemplate: StandardTemplate = {
  id: 201,
  title: t('template_excel_analyze_title'),
  desc: t('template_excel_analyze_desc'),
  endpoint: '/api/ai/excel-analyze',
  method: 'POST',
  payload: { analysisRequest: '分析excel文件中内容' }
}

function onApplyTemplate() {
  alert(t('applyTemplateAlert', { name: t('template_common_sales') }))
}

function onExportTemplates() {
  alert(t('exportTemplatesAlert'))
}

function onImportTemplates() {
  alert(t('importTemplatesAlert'))
}

async function applySalesTemplate() {
  try {
    const result = await applyTemplate(1, {})
    if (result.success) {
      alert(t('templateApplySuccess', { name: t('template_common_sales'), msg: result.message || '' }))
    }
  } catch (error) {
    console.error('应用模板失败:', error)
    alert(t('applyTemplateFailed'))
  }
}

async function applyPivotTemplate() {
  try {
    const result = await applyTemplate(2, {})
    if (result.success) {
      alert(t('templateApplySuccess', { name: t('template_pivot_dept_month'), msg: result.message || '' }))
    }
  } catch (error) {
    console.error('应用模板失败:', error)
    alert(t('applyTemplateFailed'))
  }
}

async function applyCleaningTemplate() {
  try {
    const result = await applyTemplate(3, {})
    if (result.success) {
      alert(t('templateApplySuccess', { name: t('template_cleaning'), msg: result.message || '' }))
    }
  } catch (error) {
    console.error('应用模板失败:', error)
    alert(t('applyTemplateFailed'))
  }
}

function buildPayload(tpl: StandardTemplate) {
  const merged = { ...(tpl.payload || {}) }
  const userInstr = instructionInputs.value[tpl.id]
  if (userInstr && userInstr.trim().length) merged.instructions = userInstr.trim()
  return merged
}

async function applyStandardTemplate(tpl: StandardTemplate) {
  try {
    applyingId.value = tpl.id
    const file = await getLatestSavedFile()
    if (!file) {
      alert(t('cachedFileMissing'))
      return
    }

    const base = await (await import('@/services/apiClient')).getApiBaseUrl()
    const url = tpl.endpoint.startsWith('http') ? tpl.endpoint : `${base}${tpl.endpoint.startsWith('/') ? '' : '/'}${tpl.endpoint}`
    const method = tpl.method || 'POST'

    const fd = new FormData()
    fd.append('file', file)
    Object.entries(buildPayload(tpl)).forEach(([k, v]) => fd.append(k, String(v)))

    const resp = await fetch(url, { method, body: fd })
    if (!resp.ok) {
      let body: any = null
      try { body = await resp.json() } catch (e) { }
      const msg = (body && body.error) || `HTTP ${resp.status}`
      throw new Error(msg)
    }
    const result = await resp.json()
    if (result && result.success) {
      alert(t('templateApplySuccess', { name: tpl.title, msg: result.message || '' }))
      const text = String(result.analysis || result.aiResponse || result.message || '') || JSON.stringify(result)
      emit('template-response', text)
    } else {
      alert(t('applyTemplateFailed'))
    }
  } catch (error) {
    console.error('应用模板失败:', error)
    const msg = (error as any)?.message || t('applyTemplateFailed')
    alert(msg)
    emit('template-response', msg)
  } finally {
    applyingId.value = null
  }
}

function templatePayload(tpl: StandardTemplate, minimal?: boolean) {
  const payload = buildPayload(tpl)
  const formData = minimal ? { instructions: payload.instructions || '' } : { file: '<auto-use-last-saved-file>', ...payload }
  return {
    endpoint: tpl.endpoint,
    method: tpl.method || 'POST',
    formData
  }
}

function formatJson(obj: any) {
  try { return JSON.stringify(obj, null, 2) } catch (e) { return '' }
}

async function exportTemplate(tpl: StandardTemplate) {
  try {
    const data = templatePayload(tpl, false)
    const blob = new Blob([formatJson(data)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `template_${tpl.id}.json`
    document.body.appendChild(a)
    a.click()
    a.remove()
    URL.revokeObjectURL(url)
  } catch (e) {
    alert(t('exportTemplatesAlert') || '导出失败')
  }
}

// 默认使用前端已保存的 Excel（/api/excel/save 后存储的最新文件）
async function getLatestSavedFile(): Promise<File | null> {
  try {
    const { getApiBaseUrl, fetchJson } = await import('@/services/apiClient')
    const base = await getApiBaseUrl()
    const listResp = await fetchJson('/api/excel/saved-files')
    const files = Array.isArray(listResp.files) ? listResp.files : []
    if (!files.length) return null
    // pick the latest by uploadedAt
    files.sort((a: any, b: any) => (b.uploadedAt || 0) - (a.uploadedAt || 0))
    const latest = files[0]
    const fileId = latest.fileId || latest.path || ''
    const downloadUrl = `${base}/api/excel/download?fileId=${encodeURIComponent(fileId)}`
    const blobResp = await fetch(downloadUrl)
    if (!blobResp.ok) return null
    const blob = await blobResp.blob()
    const name = latest.originalName || `saved_${fileId}.xlsx`
    return new File([blob], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  } catch (e) {
    console.warn('getLatestSavedFile failed', e)
    return null
  }
}

function loadSavedInstructions() {
  try {
    const s = localStorage.getItem('template_instruction_inputs')
    if (!s) return
    const parsed = JSON.parse(s)
    if (parsed && typeof parsed === 'object') instructionInputs.value = parsed
  } catch (e) { /* ignore */ }
}

watch(instructionInputs, (val) => {
  try {
    localStorage.setItem('template_instruction_inputs', JSON.stringify(val))
  } catch (e) { /* ignore */ }
}, { deep: true })

onMounted(() => {
  loadSavedInstructions()
})
</script>

<style scoped>
.templates {
  display: flex;
  flex-direction: column;
  height: 100%;
}

/* Ensure this page stacks rows vertically regardless of global .content */
.templates > .content {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.row-library .card {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.library-header {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
}

.library-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 12px;
}

.library-card {
  border: 1px solid rgba(255,255,255,0.05);
  border-radius: 10px;
  padding: 12px;
  background: rgba(255,255,255,0.02);
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.library-card-header {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.library-title { font-weight: 600; }

.endpoint {
  font-size: 12px;
  color: #8b5cf6;
  word-break: break-all;
}

.json-preview {
  background: #0f172a;
  color: #e2e8f0;
  padding: 8px;
  border-radius: 8px;
  font-size: 12px;
  line-height: 1.5;
  max-height: 200px;
  overflow: auto;
}

.row-second {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
}
.row-second .card { min-height: 160px; }

.row-example .card {
  display: flex;
  flex-direction: column;
  gap: 12px;
}

@media (max-width: 900px) {
  .row-second { grid-template-columns: 1fr; }
}

.template-actions {
  display: flex;
  gap: 12px;
  margin-top: 16px;
  flex-wrap: wrap;
}

.apply-template-btn,
.export-btn,
.import-btn,
.apply-btn {
  background: #7c3aed;
  color: white;
  border: 0;
  padding: 8px 16px;
  border-radius: 6px;
  cursor: pointer;
}

.apply-template-btn {
  background: #10b981;
}

.template-actions.compact {
  justify-content: flex-start;
}
</style>