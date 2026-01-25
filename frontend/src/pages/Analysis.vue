<template>
  <div class="analysis">
    <div class="topbar">
      <h2>{{ $t('analysisTitle') }}</h2>
    </div>

    <div class="content">
      <!-- 第一行：图表建议（横向平铺） -->
      <div class="row row-first">
        <div class="card charts-row">
          <div class="card-header">
            <h4>{{ $t('chartSuggestionsTitle') }}</h4>
            <p class="muted">{{ $t('chartSuggestionsDesc') }}</p>
          </div>

          <div class="data-range-row">
            <label style="font-size:12px;color:rgba(230,238,248,0.7)">数据范围</label>
            <input
              v-model="dataRangeInput"
              @input="onDataRangeInput"
              class="range-input"
              placeholder="A1:F15"
            />
            <span class="range-hint">仅输入起止单元格，例如 A1:F15（首行视为表头，首列为名称列）</span>
            <span v-if="rangeError" class="range-error">{{ rangeError }}</span>
          </div>

          <div class="chart-suggestions-grid horizontal">
            <div class="chart-suggestion-item">
              <h5>{{ $t('chart_line') }}</h5>
              <div class="chart-placeholder" style="display:flex;flex-direction:column;gap:8px;align-items:stretch;">
                <div style="font-size:12px;color:rgba(200,210,220,0.7)">数据范围：{{ dataRangeInput || '未检测' }}</div>
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
                  <canvas ref="lineCanvasRef" class="chart-canvas"></canvas>
                  <div v-if="chartInstructionsTextLine" class="chart-instructions">{{ chartInstructionsTextLine }}</div>
                  <div v-else class="chart-placeholder"></div>
                </div>
              </div>
            </div>
            <div class="chart-suggestion-item">
              <h5>{{ $t('chart_pie') }}</h5>
              <div class="chart-placeholder" style="display:flex;flex-direction:column;gap:8px;align-items:stretch;">
                <div style="font-size:12px;color:rgba(200,210,220,0.7)">数据范围：{{ dataRangeInput || '未检测' }}</div>
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
                  <canvas ref="pieCanvasRef" class="chart-canvas"></canvas>
                  <div v-if="chartInstructionsTextPie" class="chart-instructions">{{ chartInstructionsTextPie }}</div>
                  <div v-else class="chart-placeholder"></div>
                </div>
              </div>
            </div>
            <div class="chart-suggestion-item">
              <h5>{{ $t('chart_top') }}</h5>
              <div class="chart-placeholder" style="display:flex;flex-direction:column;gap:8px;align-items:stretch;">
                <div style="font-size:12px;color:rgba(200,210,220,0.7)">数据范围：{{ dataRangeInput || '未检测' }}</div>
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
                  <canvas ref="barCanvasRef" class="chart-canvas"></canvas>
                  <div v-if="chartInstructionsTextBar" class="chart-instructions">{{ chartInstructionsTextBar }}</div>
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
                {{'未生成报告：点击“生成报告”以查看预览' }}
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
import { ref, nextTick, onMounted, watch } from 'vue'
import * as XLSX from 'xlsx'
import { getExcelDataPreview } from '@/services/aiService'

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
const dataRangeInput = ref('')
const rangeError = ref('')
const selectedColumnLine = ref('')
const selectedColumnPie = ref('')
const selectedColumnBar = ref('')
const chartInstructionsTextLine = ref('')
const chartInstructionsTextPie = ref('')
const chartInstructionsTextBar = ref('')
const dataMatrix = ref<any[][]>([])
const lineCanvasRef = ref<HTMLCanvasElement | null>(null)
const pieCanvasRef = ref<HTMLCanvasElement | null>(null)
const barCanvasRef = ref<HTMLCanvasElement | null>(null)
const canvasHeight = 220
const lastLoadedFileToken = ref('')

function buildLineChartInstructions(targetColumn: string) {
  const col = targetColumn || 'Value1'
  return [
    `Certainly! Below are the step-by-step instructions to create a line chart in Excel using the "${col}" column as the target data.`,
    'Step 1: Open your Excel file and go to Sheet1.',
    `Step 2: Select the data range that includes "Name" (X-axis) and "${col}" (Y-axis). Example: A1:B15.`,
    'Step 3: Insert a 2-D Line chart via Insert > Charts > Line (first option).',
    `Step 4: Customize the chart: add a title (e.g., "${col} Over Items"), axis titles ("Items" for X, "${col}" for Y), optional data labels, choose a style, and resize/reposition as needed.`,
    'Step 5: Save the file. Optional: add a Linear trendline via Chart Design > Add Chart Element > Trendline > Linear.'
  ].join('\n')
}

function numToCol(n: number) {
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

function colToNum(col: string) {
  let n = 0
  const up = col.toUpperCase()
  for (let i = 0; i < up.length; i++) {
    const c = up.charCodeAt(i) - 64
    if (c < 1 || c > 26) return -1
    n = n * 26 + c
  }
  return n
}

function parseRange(range: string, maxRows: number, maxCols: number) {
  const trimmed = (range || '').replace(/\s+/g, '').toUpperCase()
  const m = trimmed.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
  if (!m) return null
  const sc = colToNum(m[1]) - 1
  const sr = parseInt(m[2], 10) - 1
  const ec = colToNum(m[3]) - 1
  const er = parseInt(m[4], 10) - 1
  if (sc < 0 || sr < 0 || ec < 0 || er < 0) return null
  if (sc >= maxCols || sr >= maxRows) return null
  const ecClamped = Math.min(ec, maxCols - 1)
  const erClamped = Math.min(er, maxRows - 1)
  if (sc > ecClamped || sr > erClamped) return null
  const clamped = ecClamped !== ec || erClamped !== er
  return { sc, sr, ec: ecClamped, er: erClamped, clamped }
}

function formatRange(r: { sc: number; sr: number; ec: number; er: number }) {
  return `${numToCol(r.sc + 1)}${r.sr + 1}:${numToCol(r.ec + 1)}${r.er + 1}`
}

function onDataRangeInput() {
  // keep only letters, digits, colon
  dataRangeInput.value = dataRangeInput.value.replace(/[^A-Za-z0-9:]/g, '').toUpperCase()
  rangeError.value = ''
}

function sliceByRange(matrix: any[][], r: { sc: number; sr: number; ec: number; er: number }) {
  const rows = matrix.slice(r.sr, r.er + 1)
  return rows.map(row => row.slice(r.sc, r.ec + 1))
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
        await loadMatrixFromApi(fileObj)
      }
    }
  } catch (e) {
    // ignore
  }
}

async function loadMatrixFromApi(fileObj: File) {
  try {
    const res = await getExcelDataPreview(fileObj)
    let matrix = normalizeMatrix(res)
    // if backend returns corrupted/binary-looking headers, fall back to local parse
    if (!matrix || looksCorrupted(matrix)) {
      matrix = await parseLocally(fileObj)
    }
    if (matrix) {
      dataMatrix.value = matrix
      if (matrix.length > 0 && Array.isArray(matrix[0])) {
        const headers = (matrix[0] as any[]).map(c => (c == null ? '' : String(c).trim()))
        sheetHeaders.value = headers
        const firstValueHeader = headers.slice(1).find(h => h && h.trim().length > 0) || headers[1] || ''
        selectedColumnLine.value = firstValueHeader
        selectedColumnPie.value = firstValueHeader
        selectedColumnBar.value = firstValueHeader
        const lastRow = matrix.length
        const lastCol = headers.length || 1
        dataRange.value = `A1:${numToCol(lastCol)}${lastRow}`
        dataRangeInput.value = dataRange.value
      }
    }
  } catch (e) {
    rangeError.value = '无法从后端获取数据，请检查文件'
  }
}

function looksCorrupted(matrix: any[][]) {
  if (!matrix || !matrix.length) return true
  const headerRow = matrix[0]
  if (!Array.isArray(headerRow)) return true
  const joined = headerRow.map(c => (c == null ? '' : String(c))).join(' ')
  return /PK\w+workbook|_rels\/workbook\.xml\.rels/i.test(joined)
}

async function parseLocally(fileObj: File): Promise<any[][] | null> {
  try {
    const buf = await fileObj.arrayBuffer()
    const wb = XLSX.read(buf, { type: 'array' })
    const first = wb.SheetNames[0]
    const sheet = wb.Sheets[first]
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][]
    if (Array.isArray(aoa) && aoa.length) return aoa
  } catch (err) {
    // ignore
  }
  return null
}

function normalizeMatrix(res: any): any[][] | null {
  // preferred: res.data is array of arrays
  if (res && Array.isArray(res.data) && res.data.every((r: any) => Array.isArray(r))) return res.data as any[][]
  // if data is a string, try JSON parse
  if (res && typeof res.data === 'string') {
    try {
      const parsed = JSON.parse(res.data)
      if (Array.isArray(parsed) && parsed.every((r: any) => Array.isArray(r))) return parsed as any[][]
    } catch (err) {
      /* ignore */
    }
  }
  // try excelDataPreview string in TSV/CSV-ish format
  const preview = res && typeof res.excelDataPreview === 'string' ? res.excelDataPreview : null
  if (preview) {
    const lines = preview.split(/\r?\n/).filter((l: string) => l.trim().length > 0)
    const rows = lines.map((l: string) => l.split(/\t|,/))
    if (rows.length) return rows
  }
  return null
}

// draw helpers
function resizeCanvas(canvas: HTMLCanvasElement) {
  const w = canvas.clientWidth || 360
  const h = canvas.clientHeight || canvasHeight
  if (canvas.width !== w) canvas.width = w
  if (canvas.height !== h) canvas.height = h
}

function clearCanvas(canvas: HTMLCanvasElement) {
  resizeCanvas(canvas)
  const ctx = canvas.getContext('2d')
  if (!ctx) return
  ctx.clearRect(0, 0, canvas.width, canvas.height)
  ctx.fillStyle = 'rgba(255,255,255,0.02)'
  ctx.fillRect(0, 0, canvas.width, canvas.height)
}

function drawLine(canvas: HTMLCanvasElement, labels: string[], values: number[]) {
  const ctx = canvas.getContext('2d')
  if (!ctx) return
  clearCanvas(canvas)
  const w = canvas.width - 60
  const h = canvas.height - 60
  const ox = 40
  const oy = canvas.height - 30
  const max = Math.max(...values)
  const min = Math.min(...values)
  const span = Math.max(max - min, 1)
  const step = Math.max(1, Math.ceil(span / Math.max(values.length, 1)))
  const yStart = Math.floor(min)

  // gridlines and y-axis ticks
  ctx.strokeStyle = 'rgba(255,255,255,0.15)'
  ctx.fillStyle = 'rgba(255,255,255,0.5)'
  ctx.font = '10px sans-serif'
  for (let yv = yStart; yv <= max + step; yv += step) {
    const yPos = oy - ((yv - min) / span) * h
    ctx.beginPath()
    ctx.moveTo(ox, yPos)
    ctx.lineTo(ox + w, yPos)
    ctx.stroke()
    ctx.fillText(String(yv), 4, yPos + 3)
  }

  ctx.strokeStyle = 'rgba(120,179,255,0.8)'
  ctx.beginPath()
  values.forEach((v, i) => {
    const x = ox + (w * i) / Math.max(values.length - 1, 1)
    const y = oy - ((v - min) / span) * h
    if (i === 0) ctx.moveTo(x, y)
    else ctx.lineTo(x, y)
  })
  ctx.stroke()
  ctx.fillStyle = 'rgba(120,179,255,0.8)'
  values.forEach((v, i) => {
    const x = ox + (w * i) / Math.max(values.length - 1, 1)
    const y = oy - ((v - min) / span) * h
    ctx.beginPath(); ctx.arc(x, y, 3, 0, Math.PI * 2); ctx.fill()
  })
  // axes
  ctx.strokeStyle = 'rgba(255,255,255,0.2)'
  ctx.beginPath(); ctx.moveTo(ox, oy - h); ctx.lineTo(ox, oy); ctx.lineTo(ox + w, oy); ctx.stroke()
  ctx.fillStyle = 'rgba(255,255,255,0.5)'
  ctx.font = '10px sans-serif'
  labels.forEach((l: string, i: number) => {
    const x = ox + (w * i) / Math.max(labels.length - 1, 1)
    ctx.fillText(l, x - 10, oy + 12)
  })
}

function drawBar(canvas: HTMLCanvasElement, labels: string[], values: number[]) {
  const ctx = canvas.getContext('2d')
  if (!ctx) return
  clearCanvas(canvas)
  const w = canvas.width - 60
  const h = canvas.height - 60
  const ox = 40
  const oy = canvas.height - 30
  const max = Math.max(...values)
  const min = Math.min(...values)
  const span = Math.max(max - min, 1)
  const step = Math.max(1, Math.ceil(span / Math.max(values.length, 1)))
  const barW = w / Math.max(values.length, 1) * 0.6

  // gridlines and y-axis ticks
  ctx.strokeStyle = 'rgba(255,255,255,0.15)'
  ctx.fillStyle = 'rgba(255,255,255,0.5)'
  ctx.font = '10px sans-serif'
  for (let yv = Math.floor(min); yv <= max + step; yv += step) {
    const yPos = oy - ((yv - min) / span) * h
    ctx.beginPath(); ctx.moveTo(ox, yPos); ctx.lineTo(ox + w, yPos); ctx.stroke()
    ctx.fillText(String(yv), 4, yPos + 3)
  }

  ctx.fillStyle = 'rgba(125,86,255,0.8)'
  values.forEach((v, i) => {
    const x = ox + (w / Math.max(values.length, 1)) * i + (w / Math.max(values.length, 1) - barW) / 2
    const y = oy - ((v - min) / span) * h
    ctx.fillRect(x, y, barW, oy - y)
  })
  ctx.strokeStyle = 'rgba(255,255,255,0.2)'
  ctx.beginPath(); ctx.moveTo(ox, oy - h); ctx.lineTo(ox, oy); ctx.lineTo(ox + w, oy); ctx.stroke()
  ctx.fillStyle = 'rgba(255,255,255,0.5)'
  ctx.font = '10px sans-serif'
  labels.forEach((l: string, i: number) => {
    const x = ox + (w / Math.max(values.length, 1)) * i
    ctx.fillText(l, x, oy + 12)
  })
}

function drawPie(canvas: HTMLCanvasElement, labels: string[], values: number[]) {
  const ctx = canvas.getContext('2d')
  if (!ctx) return
  clearCanvas(canvas)
  const total = values.reduce((a, b) => a + b, 0)
  if (total <= 0) return
  let start = -Math.PI / 2
  const cx = canvas.width / 2
  const cy = canvas.height / 2
  const r = Math.min(canvas.width, canvas.height) / 3
  values.forEach((v, i) => {
    const pct = v / total
    const end = start + pct * Math.PI * 2
    const hue = (i * 60) % 360
    ctx.beginPath()
    ctx.moveTo(cx, cy)
    ctx.fillStyle = `hsl(${hue},70%,60%)`
    ctx.arc(cx, cy, r, start, end)
    ctx.closePath()
    ctx.fill()
    start = end
  })
  ctx.fillStyle = 'rgba(255,255,255,0.85)'
  ctx.font = '10px sans-serif'
  let acc = -Math.PI / 2
  values.forEach((v, i) => {
    const pct = v / total
    const mid = acc + pct * Math.PI
    const rx = cx + Math.cos(mid) * (r + 14)
    const ry = cy + Math.sin(mid) * (r + 14)
    const label = labels[i] || `项${i+1}`
    const percent = `${(pct * 100).toFixed(1)}%`
    ctx.fillText(`${label} ${percent}`, rx - 18, ry)
    acc += pct * Math.PI * 2
  })
}

function extractSeries(matrix: any[][], targetColumn: string, range: { sc: number; sr: number; ec: number; er: number }) {
  const sliced = sliceByRange(matrix, range)
  if (sliced.length === 0) return { labels: [], values: [] }
  const headerRow = sliced[0]
  const targetIdx = headerRow.findIndex((h: any) => (h == null ? '' : String(h).trim()) === (targetColumn || '').trim())
  if (targetIdx <= 0) return { labels: [], values: [] }
  const labels: string[] = []
  const values: number[] = []
  for (let i = 1; i < sliced.length; i++) {
    const row = sliced[i]
    const name = row[0]
    const v = Number(row[targetIdx])
    if (!isFinite(v)) continue
    labels.push(name == null ? `项${i}` : String(name))
    values.push(v)
  }
  return { labels, values }
}

async function onCreateChart(chartType: string, targetColumn: string) {
  if (!savedFile.value) { alert('未检测到服务器缓存的 Excel 文件'); return }
  rangeError.value = ''
  if (!targetColumn) {
    if (chartType === 'line') chartInstructionsTextLine.value = '请选择目标列'
    if (chartType === 'pie') chartInstructionsTextPie.value = '请选择目标列'
    if (chartType === 'bar') chartInstructionsTextBar.value = '请选择目标列'
    return
  }

  if (!dataMatrix.value.length) {
    const msg = '未检测到可用数据，请先上传文件并确认范围'
    if (chartType === 'line') chartInstructionsTextLine.value = msg
    if (chartType === 'pie') chartInstructionsTextPie.value = msg
    if (chartType === 'bar') chartInstructionsTextBar.value = msg
    return
  }

  const parsed = parseRange(dataRangeInput.value || dataRange.value, dataMatrix.value.length, dataMatrix.value[0]?.length || 0)
  if (!parsed) {
    rangeError.value = '范围格式无效或超出数据大小'
    return
  }
  if ((parsed as any).clamped) {
    const safeRange = formatRange(parsed)
    dataRangeInput.value = safeRange
    rangeError.value = `输入范围超出数据，已裁剪为 ${safeRange}`
  }

  const { labels, values } = extractSeries(dataMatrix.value, targetColumn, parsed)
  if (!labels.length || !values.length) {
    const msg = '未找到有效数据，请确认首列为名称列，目标列为数值列'
    if (chartType === 'line') chartInstructionsTextLine.value = msg
    if (chartType === 'pie') chartInstructionsTextPie.value = msg
    if (chartType === 'bar') chartInstructionsTextBar.value = msg
    return
  }

  if (chartType === 'line' && lineCanvasRef.value) drawLine(lineCanvasRef.value, labels, values)
  if (chartType === 'pie' && pieCanvasRef.value) drawPie(pieCanvasRef.value, labels, values)
  if (chartType === 'bar' && barCanvasRef.value) drawBar(barCanvasRef.value, labels, values)

  const successMsg = chartType === 'line'
    ? `已基于列 "${targetColumn}" 生成${chartType === 'line'}折线图（首行表头，首列名称，已过滤非数值行）`
    : `已基于列 "${targetColumn}" 生成${chartType === 'pie' ? '扇形图' : '柱状图'}（首行表头，首列名称，已过滤非数值行）`
  if (chartType === 'line') chartInstructionsTextLine.value = successMsg
  if (chartType === 'pie') chartInstructionsTextPie.value = successMsg
  if (chartType === 'bar') chartInstructionsTextBar.value = successMsg
}

onMounted(() => {
  loadLatestSavedFile()
})

watch(savedFile, async (f) => {
  if (!f) return
  const token = `${f.name}:${f.size}`
  if (token === lastLoadedFileToken.value) return
  lastLoadedFileToken.value = token
  await loadMatrixFromApi(f)
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

.chart-instructions {
  white-space: pre-line;
  line-height: 1.4;
}

.data-range-row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin: 8px 0 0 0;
}

.range-input {
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid rgba(255, 255, 255, 0.08);
  color: #fff;
  padding: 6px 8px;
  border-radius: 4px;
  width: 140px;
}

.range-hint {
  font-size: 12px;
  color: rgba(230, 238, 248, 0.6);
}

.range-error {
  font-size: 12px;
  color: #ffb4b4;
}

.chart-canvas {
  width: 100%;
  height: 220px;
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid rgba(255, 255, 255, 0.04);
  border-radius: 4px;
  margin-bottom: 6px;
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