<template>
  <div class="analysis">
    <div class="topbar">
      <h2>{{ $t('analysisTitle') }}</h2>
      <div class="muted">{{ $t('analysisSubtitle') }}</div>
    </div>

    <div class="content">
      <!-- 第一行：图表建议（横向平铺） -->
      <div class="row row-first">
        <div class="card charts-row">
          <div class="card-header">
            <h4>{{ $t('chartSuggestionsTitle') }}</h4>
            <p class="muted">{{ $t('chartSuggestionsDesc') }}</p>
          </div>

          <div class="chart-suggestions-grid horizontal">
            <div class="chart-suggestion-item">
              <h5>{{ $t('chart_line') }}</h5>
              <div class="chart-placeholder" style="display:flex;flex-direction:column;gap:8px;align-items:stretch;">
                <div style="font-size:12px;color:rgba(200,210,220,0.7)">数据范围：{{ dataRange || '未检测' }}</div>
                <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
                  <label style="font-size:12px;color:rgba(230,238,248,0.7)">目标列</label>
                  <select v-model="selectedColumnLine">
                    <option v-for="(h,idx) in sheetHeaders" :key="idx" :value="h">{{ h || ('列 ' + (idx+1)) }}</option>
                  </select>
                </div>
                <div style="display:flex;gap:8px;">
                  <button @click="onCreateChart('line', selectedColumnLine)" class="generate-report-btn" :disabled="!savedFile">创建图表</button>
                </div>
                <div style="margin-top:6px;color:rgba(230,238,248,0.9);font-size:13px;">
                  <div v-if="chartInstructionsTextLine">{{ chartInstructionsTextLine }}</div>
                  <div v-else class="chart-placeholder"></div>
                </div>
              </div>
            </div>
            <div class="chart-suggestion-item">
              <h5>{{ $t('chart_pie') }}</h5>
              <div class="chart-placeholder" style="display:flex;flex-direction:column;gap:8px;align-items:stretch;">
                <div style="font-size:12px;color:rgba(200,210,220,0.7)">数据范围：{{ dataRange || '未检测' }}</div>
                <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
                  <label style="font-size:12px;color:rgba(230,238,248,0.7)">目标列</label>
                  <select v-model="selectedColumnPie">
                    <option v-for="(h,idx) in sheetHeaders" :key="idx" :value="h">{{ h || ('列 ' + (idx+1)) }}</option>
                  </select>
                </div>
                <div style="display:flex;gap:8px;">
                  <button @click="onCreateChart('pie', selectedColumnPie)" class="generate-report-btn" :disabled="!savedFile">创建图表</button>
                </div>
                <div style="margin-top:6px;color:rgba(230,238,248,0.9);font-size:13px;">
                  <div v-if="chartInstructionsTextPie">{{ chartInstructionsTextPie }}</div>
                  <div v-else class="chart-placeholder"></div>
                </div>
              </div>
            </div>
            <div class="chart-suggestion-item">
              <h5>{{ $t('chart_top') }}</h5>
              <div class="chart-placeholder" style="display:flex;flex-direction:column;gap:8px;align-items:stretch;">
                <div style="font-size:12px;color:rgba(200,210,220,0.7)">数据范围：{{ dataRange || '未检测' }}</div>
                <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
                  <label style="font-size:12px;color:rgba(230,238,248,0.7)">目标列</label>
                  <select v-model="selectedColumnBar">
                    <option v-for="(h,idx) in sheetHeaders" :key="idx" :value="h">{{ h || ('列 ' + (idx+1)) }}</option>
                  </select>
                </div>
                <div style="display:flex;gap:8px;">
                  <button @click="onCreateChart('bar', selectedColumnBar)" class="generate-report-btn" :disabled="!savedFile">创建图表</button>
                </div>
                <div style="margin-top:6px;color:rgba(230,238,248,0.9);font-size:13px;">
                  <div v-if="chartInstructionsTextBar">{{ chartInstructionsTextBar }}</div>
                  <div v-else class="chart-placeholder"></div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- 第二行：自动报告（占满宽度，高度720px） -->
        <div class="row row-second">
          <div class="card auto-report-card" :class="{ generated: reportGenerated }">
            <h4>{{ $t('autoReportTitle') }}</h4>
            <p class="muted">{{ $t('autoReportDesc') }}</p>
            <button @click="onGenerateReportClick" class="generate-report-btn">{{ $t('generateReportBtn') }}</button>
            <button @click="onLoadApiExample" class="generate-report-btn" style="margin-left:8px">{{ $t('loadApiExample') || '加载接口示例' }}</button>

            <!-- 报告主体：未生成时高度 160px；生成后根据内容测量高度；当内容高度 > 720px 时启用滚动（鼠标滚轮预览） -->
            <div class="auto-report-body" :style="{ height: (reportGenerated ? reportHeight + 'px' : '160px'), overflowY: reportHeight > 720 ? 'auto' : 'hidden' }">
              <div v-if="!reportGenerated" class="report-placeholder">
                {{ $t('autoReportPlaceholder') || '未生成报告：点击“生成报告”以查看预览' }}
              </div>
              <div v-else class="report-content" v-html="reportHtml"></div>
            </div>
          </div>
        </div>

      <!-- 第三行：RFM / CLV / 财务，三列等分 -->
      <div class="row row-third">
        <div class="card">
          <h4>{{ $t('rfmTitle') }}</h4>
          <p class="muted">{{ $t('rfmDesc') }}</p>
          <button @click="onRfmClick" class="rfm-btn">{{ $t('rfmBtn') }}</button>
        </div>

        <div class="card">
          <h4>{{ $t('clvTitle') }}</h4>
          <p class="muted">{{ $t('clvDesc') }}</p>
          <button @click="onClvClick" class="clv-btn">{{ $t('clvBtn') }}</button>
        </div>

        <div class="card">
          <h4>{{ $t('financeTitle') }}</h4>
          <p class="muted">{{ $t('financeDesc') }}</p>
          <button @click="onFinanceClick" class="finance-btn">{{ $t('financeBtn') }}</button>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import { getAnalysisCenterData } from '@/services/api'
import { ref, nextTick, onMounted } from 'vue'
import * as XLSX from 'xlsx'
import { createChart } from '@/services/aiService'

const { t } = useI18n()

const reportGenerated = ref(false)
const reportGenerating = ref(false)
const reportHeight = ref(160)
const reportHtml = ref('')
// Chart creation state (use server-cached file)
const savedFile = ref<File | null>(null)
const savedFileId = ref<string | null>(null)
const sheetHeaders = ref<string[]>([])
const dataRange = ref('')
const selectedColumnLine = ref('')
const selectedColumnPie = ref('')
const selectedColumnBar = ref('')
const chartInstructionsTextLine = ref('')
const chartInstructionsTextPie = ref('')
const chartInstructionsTextBar = ref('')

function numToCol(n: number) {
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

// load latest saved file metadata and file bytes from backend
async function loadLatestSavedFile() {
  try {
    const resp = await fetch('/api/excel/saved-files')
    if (!resp.ok) return
    const j = await resp.json()
    if (j && Array.isArray(j.files) && j.files.length > 0) {
      // pick the last entry as the most recent
      const last = j.files[j.files.length - 1]
      if (last && last.fileId) {
        savedFileId.value = last.fileId
        // download file bytes
        const dl = await fetch(`/api/excel/download?fileId=${encodeURIComponent(last.fileId)}`)
        if (!dl.ok) return
        const blob = await dl.blob()
        // try to construct File with originalName if available
        const filename = last.originalName || last.fileId || 'saved.xlsx'
        const fileObj = new File([blob], filename, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
        savedFile.value = fileObj
        // parse headers and data range
        try {
          const arrayBuf = await blob.arrayBuffer()
          const wb = XLSX.read(arrayBuf, { type: 'array' })
          const first = wb.SheetNames[0]
          const sheet = wb.Sheets[first]
          const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[]
          if (aoa && aoa.length > 0) {
            const headers = aoa[0].map((c:any) => c == null ? '' : String(c))
            sheetHeaders.value = headers
            selectedColumnLine.value = headers[0] || ''
            selectedColumnPie.value = headers[0] || ''
            selectedColumnBar.value = headers[0] || ''
            const lastRow = aoa.length
            const lastCol = headers.length || 1
            dataRange.value = `A1:${numToCol(lastCol)}${lastRow}`
          }
        } catch (e) {
          // ignore parse errors
        }
      }
    }
  } catch (e) {
    // ignore
  }
}

// create chart using saved file or savedFileId
async function onCreateChart(chartType: string, targetColumn: string) {
  if (!savedFile.value && !savedFileId.value) { alert('未检测到服务器缓存的 Excel 文件'); return }
  try {
    // set appropriate status text per chart type
    if (chartType === 'line') chartInstructionsTextLine.value = '正在创建图表...'
    if (chartType === 'pie') chartInstructionsTextPie.value = '正在创建图表...'
    if (chartType === 'bar') chartInstructionsTextBar.value = '正在创建图表...'

    const res = await createChart(savedFile.value || null, chartType, targetColumn, savedFileId.value || undefined)

    const instr = res && (res as any).chartInstructions ? (res as any).chartInstructions : JSON.stringify(res || {})
    if (chartType === 'line') chartInstructionsTextLine.value = instr
    if (chartType === 'pie') chartInstructionsTextPie.value = instr
    if (chartType === 'bar') chartInstructionsTextBar.value = instr
  } catch (e:any) {
    const msg = (e && e.message) || '创建图表失败'
    if (chartType === 'line') chartInstructionsTextLine.value = msg
    if (chartType === 'pie') chartInstructionsTextPie.value = msg
    if (chartType === 'bar') chartInstructionsTextBar.value = msg
  }
}

onMounted(() => {
  loadLatestSavedFile()
})

async function onRecommendClick() {
  try {
    const result = await getAnalysisCenterData()
    if (result.success) {
      alert(t('recommendAlert'))
    }
  } catch (error) {
    console.error('获取推荐失败:', error)
    alert('获取推荐失败')
  }
}

async function onGenerateReportClick() {
  // 标记为生成中，延迟 2 秒以显示“生成中”状态
  reportGenerating.value = true
  await new Promise(resolve => setTimeout(resolve, 2000))

  // 尝试从 API 获取报告内容；若失败则使用示例占位内容
  try {
    const result = await getAnalysisCenterData()
    // 兼容两种返回结构：{ content: string } 或 { data: { ... } }
    if (result && (result as any).content) {
      reportHtml.value = (result as any).content
    } else if (result && (result as any).data) {
      const d = (result as any).data
      const topProducts = Array.isArray(d.topProducts) ? d.topProducts.map((p: any) => `<li>${p}</li>`).join('') : ''
      const cust = d.customerDistribution || {}
      reportHtml.value = `
        <h3>销售趋势</h3>
        <p>${d.salesTrend || ''}</p>
        <h3>畅销产品</h3>
        <ul>${topProducts}</ul>
        <h3>客户分布</h3>
        <p>新客户: ${cust['新客户'] || 0}, 老客户: ${cust['老客户'] || 0}</p>
        <h3>营收增长</h3>
        <p>${d.revenueGrowth || ''}</p>
      `
    } else {
      reportHtml.value = '<p>' + t('generateReportAlert') + '</p>'
    }
  } catch (e) {
    // 生成较长的示例内容以演示滚动
    reportHtml.value = Array.from({ length: 40 }).map((_, i) => `<p>示例报告段落 ${i+1}：这是一段示例文本，用于测试自动报告的显示与滚动。</p>`).join('')
  }

  reportGenerated.value = true
  reportGenerating.value = false
  // 等待 DOM 更新然后测量实际高度
  await nextTick()
  const el = document.querySelector('.auto-report-card .report-content') as HTMLElement | null
  if (el) {
    // 使用 scrollHeight 作为内容高度
    reportHeight.value = el.scrollHeight
  } else {
    reportHeight.value = 800
  }
}

function onRfmClick() {
  alert(t('rfmBtn') + ' clicked')
}

// 从内置 API 文档示例加载一个响应示例到报告预览
function onLoadApiExample() {
  const example = {
    success: true,
    analysis: "根据提供的销售数据，我发现以下趋势：1.销售量在前3个月持续增长，2.随后在第4-6个月出现下降，3.第7-9个月开始恢复增长，4.最后3个月达到年度高峰。建议在销售淡季加强营销活动。",
    excelDataPreview: "Sheet: Sales\\nMonth\\tSales\\tRegion\\t...\\n...",
    analysisRequest: "分析销售趋势",
    commandResults: []
  }

  reportHtml.value = `<pre><code>${JSON.stringify(example, null, 2)}</code></pre>`
  reportGenerated.value = true
  // 等待 DOM 更新然后测量高度
  nextTick().then(() => {
    const el = document.querySelector('.auto-report-card .report-content') as HTMLElement | null
    if (el) reportHeight.value = el.scrollHeight
    else reportHeight.value = 800
  })
}

function onClvClick() {
  alert(t('clvBtn') + ' clicked')
}

function onFinanceClick() {
  alert(t('financeBtn') + ' clicked')
}
</script>

<style scoped>
.analysis {
  display: grid;
  grid-template-rows: auto 1fr;
  gap: 16px;
  height: 100%;
}

.analysis > .topbar { grid-row: 1; }
.analysis > .content { grid-row: 2; }

/* 覆盖全局 .content 的横向 flex 布局，强制本页面垂直堆叠行 */
.analysis > .content {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.chart-suggestions-grid {
  display: grid;
  gap: 16px;
  margin: 16px 0;
}

.chart-suggestion-item {
  border: 1px solid rgba(255, 255, 255, 0.06);
  border-radius: 6px;
  padding: 12px;
}

.chart-placeholder {
  height: 120px;
  background: rgba(255, 255, 255, 0.02);
  border-radius: 4px;
  display: flex;
  align-items: flex-start;
  justify-content: flex-start;
  margin-top: 0;
  color: rgba(230, 238, 248, 0.6);
  font-size: 14px;
}

.recommend-btn,
.generate-report-btn,
.rfm-btn,
.clv-btn,
.finance-btn {
  background: #7c3aed;
  color: white;
  border: 0;
  padding: 8px 16px;
  border-radius: 6px;
  cursor: pointer;
  margin-top: 12px;
}

/* Two-row layout */
.row {
  width: 100%;
  display: grid;
  gap: 16px;
}
.row-first {
  grid-template-columns: 1fr;
}
.charts-row { min-height: 600px; height: auto; }
.chart-suggestions-grid.horizontal {
  display: flex;
  gap: 16px;
  height: auto; /* let content determine height to avoid large bottom gaps */
}
.chart-suggestion-item {
  flex: 1 1 0;
  display: flex;
  flex-direction: column;
}
.chart-suggestion-item .chart-placeholder {
  /* increase visual area so there is less empty space above/below controls */
  flex: 0 0 70%;
  height: 70%;
  min-height: 260px;
  margin-top: 0;
}

/* tighten card/header spacing to remove top gap */
.card {
  padding: 8px;
}
.card-header { padding: 8px 8px 4px 8px; }
.card-header h4 { margin: 0; font-size: 16px; }
.card-header .muted { margin-top: 4px; font-size: 12px; }

/* ensure card internals stack tightly */
.chart-suggestion-item { padding: 8px; gap: 8px; justify-content: flex-start; }

.row-second { grid-template-columns: 1fr; }
.auto-report-card { min-height: 160px; }
.auto-report-card .auto-report-body { transition: height 200ms ease; }
.auto-report-card .report-placeholder { display:flex; align-items:center; justify-content:center; height:100%; color: rgba(230,238,248,0.6); }
.auto-report-card .report-content { padding:12px; color: rgba(230,238,248,0.9); }
.auto-report-card.generated { /* extra visual emphasis when generated */ box-shadow: 0 2px 10px rgba(0,0,0,0.3); }

.row-third {
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
  margin-top: 16px;
}
.row-third .card { min-height: 160px; }

@media (max-width: 900px) {
  .row-first, .row-second {
    grid-template-columns: 1fr;
  }
  .row-first .card { min-height: 280px; }
  .row-second .card { min-height: 140px; }
}
</style>