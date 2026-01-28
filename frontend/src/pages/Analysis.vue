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
            <label style="font-size:12px;color:rgba(230,238,248,0.7)">{{ $t('dataRangeLabel') }}</label>
            <input
              v-model="dataRangeInput"
              @input="onDataRangeInput"
              class="range-input"
              :placeholder="$t('dataRangePlaceholder')"
            />
            <span class="range-hint">{{ $t('rangeHint') }}</span>
            <span v-if="rangeError" class="range-error">{{ rangeError }}</span>
          </div>

          <div class="chart-suggestions-grid horizontal">
            <div class="chart-suggestion-item">
              <h5>{{ $t('chart_line') }}</h5>
              <div class="chart-placeholder" style="display:flex;flex-direction:column;gap:8px;align-items:stretch;">
                <div style="font-size:12px;color:rgba(200,210,220,0.7)">{{ $t('dataRangeLabel') }}：{{ dataRangeInput || $t('rangeNotDetected') }}</div>
                <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
                  <label style="font-size:12px;color:rgba(230,238,248,0.7)">{{ $t('targetColumnLabel') }}</label>
                  <select v-model="selectedColumnLine">
                    <option v-for="(h,idx) in sheetHeaders" :key="idx" :value="h">{{ h || $t('col_default', { n: idx + 1 }) }}</option>
                  </select>
                </div>
                <div style="display:flex;gap:8px;">
                  <button @click="onCreateChart('line', selectedColumnLine)" class="generate-report-btn" :disabled="!savedFile">{{ $t('createChartBtn') }}</button>
                </div>
                <div style="margin-top:6px;color:rgba(230,238,248,0.9);font-size:13px;">
                  <canvas ref="lineCanvasRef" class="chart-canvas"></canvas>
                  <div v-if="chartInstructionsTextLine" class="chart-instructions">{{ chartInstructionsTextLine }}</div>
                </div>
              </div>
            </div>
            <div class="chart-suggestion-item">
              <h5>{{ $t('chart_pie') }}</h5>
              <div class="chart-placeholder" style="display:flex;flex-direction:column;gap:8px;align-items:stretch;">
                <div style="font-size:12px;color:rgba(200,210,220,0.7)">{{ $t('dataRangeLabel') }}：{{ dataRangeInput || $t('rangeNotDetected') }}</div>
                <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
                  <label style="font-size:12px;color:rgba(230,238,248,0.7)">{{ $t('targetColumnLabel') }}</label>
                  <select v-model="selectedColumnPie">
                    <option v-for="(h,idx) in sheetHeaders" :key="idx" :value="h">{{ h || $t('col_default', { n: idx + 1 }) }}</option>
                  </select>
                </div>
                <div style="display:flex;gap:8px;">
                  <button @click="onCreateChart('pie', selectedColumnPie)" class="generate-report-btn" :disabled="!savedFile">{{ $t('createChartBtn') }}</button>
                </div>
                <div style="margin-top:6px;color:rgba(230,238,248,0.9);font-size:13px;">
                  <canvas ref="pieCanvasRef" class="chart-canvas"></canvas>
                  <div v-if="chartInstructionsTextPie" class="chart-instructions">{{ chartInstructionsTextPie }}</div>
                </div>
              </div>
            </div>
            <div class="chart-suggestion-item">
              <h5>{{ $t('chart_top') }}</h5>
              <div class="chart-placeholder" style="display:flex;flex-direction:column;gap:8px;align-items:stretch;">
                <div style="font-size:12px;color:rgba(200,210,220,0.7)">{{ $t('dataRangeLabel') }}：{{ dataRangeInput || $t('rangeNotDetected') }}</div>
                <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
                  <label style="font-size:12px;color:rgba(230,238,248,0.7)">{{ $t('targetColumnLabel') }}</label>
                  <select v-model="selectedColumnBar">
                    <option v-for="(h,idx) in sheetHeaders" :key="idx" :value="h">{{ h || $t('col_default', { n: idx + 1 }) }}</option>
                  </select>
                </div>
                <div style="display:flex;gap:8px;">
                  <button @click="onCreateChart('bar', selectedColumnBar)" class="generate-report-btn" :disabled="!savedFile">{{ $t('createChartBtn') }}</button>
                </div>
                <div style="margin-top:6px;color:rgba(230,238,248,0.9);font-size:13px;">
                  <canvas ref="barCanvasRef" class="chart-canvas"></canvas>
                  <div v-if="chartInstructionsTextBar" class="chart-instructions">{{ chartInstructionsTextBar }}</div>
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
            <div class="report-actions">
              <div class="report-buttons">
                <button @click="onGenerateReportClick" class="generate-report-btn">{{ $t('generateReportBtn') }}</button>
                <button @click="onLoadApiExample" class="generate-report-btn" style="margin-left:8px">{{ $t('loadApiExample') }}</button>
                <button
                  v-if="reportGenerated"
                  @click="onDownloadReport"
                  class="generate-report-btn download-report-btn"
                >下载报告</button>
              </div>
              <div class="report-options">
                <span class="options-label">{{ $t('optionalAnalyses') }}</span>
                <label><input type="checkbox" v-model="includeFinancialRatios" />{{ $t('includeFinancialRatios') }}</label>
                <label><input type="checkbox" v-model="includeProfitability" />{{ $t('includeProfitability') }}</label>
                <label><input type="checkbox" v-model="includeCashFlow" />{{ $t('includeCashFlow') }}</label>
                <label><input type="checkbox" v-model="includeBudgetActual" />{{ $t('includeBudgetActual') }}</label>
                <label><input type="checkbox" v-model="includeRfm" />{{ $t('includeRfm') }}</label>
                <label><input type="checkbox" v-model="includeClv" />{{ $t('includeClv') }}</label>
              </div>
            </div>

            <!-- 报告主体：未生成时高度 160px；生成后根据内容测量高度；当内容高度 > 720px 时启用滚动（鼠标滚轮预览） -->
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
                <XMarkdown :markdown="reportMarkdown" :highlight="true" />
              </div>
            </div>
          </div>
        </div>

      <!-- 第三行：RFM / CLV / 财务，三列等分 -->
      <!-- <div class="row row-third">
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
      </div> -->
    </div>
  </div>
</template>

<script setup lang="ts">
import { useI18n } from 'vue-i18n'
import { getAnalysisCenterData } from '@/services/api'
import { ref, nextTick, onMounted, watch } from 'vue'
import { XMarkdown } from 'vue-element-plus-x'
import * as XLSX from 'xlsx'
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import {
  analyzeBudgetActual,
  analyzeCashFlow,
  analyzeClv,
  analyzeFinancial,
  analyzeFinancialRatios,
  analyzeProfitability,
  analyzeRfm,
  getExcelDataPreview
} from '@/services/aiService'

const { t } = useI18n()

const reportGenerated = ref(false)
const reportGenerating = ref(false)
const reportHeight = ref(160)
const reportMarkdown = ref('')
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
const includeFinancialRatios = ref(false)
const includeProfitability = ref(false)
const includeCashFlow = ref(false)
const includeBudgetActual = ref(false)
const includeRfm = ref(false)
const includeClv = ref(false)

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
    rangeError.value = t('fetchDataFailed')
  }
}

function looksCorrupted(matrix: any[][]) {
  if (!matrix || !matrix.length) return true
  const headerRow = matrix[0]
  if (!Array.isArray(headerRow)) return true
  const joined = headerRow.map(c => (c == null ? '' : String(c))).join(' ')
  return /PK\w+workbook|_rels\/workbook\.xml\.rels/i.test(joined)
}

function headerLooksMojibake(row: any[]) {
  const joined = (row || []).map(c => (c == null ? '' : String(c))).join(' ')
  if (!joined) return true
  if (joined.includes('\ufffd') || joined.includes('�')) return true
  return /锟|Ã|Â|ä¸|å|æ|ç|é¡|äº|å½/.test(joined)
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

function renderSectionMd(title: string, body?: string, preview?: string) {
  const safeBody = (body && body.trim()) ? body.trim() : t('reportNoContent')
  const previewBlock = preview ? `\n\n**${t('reportPreviewTitle')}**\n\n\u0060\u0060\u0060\n${preview}\n\u0060\u0060\u0060` : ''
  return `### ${title}\n\n${safeBody}${previewBlock}`
}

function renderErrorMd(title: string, err: unknown) {
  const msg = (err as any)?.message || String(err) || t('reportUnknownError')
  const body = t('reportApiError', { name: title, msg })
  return `### ${title}\n\n${body}\n\n> ${msg}`
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
    const label = labels[i] || t('itemLabel', { n: i + 1 })
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
    labels.push(name == null ? t('itemLabel', { n: i }) : String(name))
    values.push(v)
  }
  return { labels, values }
}

async function onCreateChart(chartType: string, targetColumn: string) {
  if (!savedFile.value) { alert(t('cachedFileMissing')); return }
  rangeError.value = ''
  if (!targetColumn) {
    if (chartType === 'line') chartInstructionsTextLine.value = t('selectTargetColumn')
    if (chartType === 'pie') chartInstructionsTextPie.value = t('selectTargetColumn')
    if (chartType === 'bar') chartInstructionsTextBar.value = t('selectTargetColumn')
    return
  }

  if (!dataMatrix.value.length) {
    const msg = t('noDataAvailable')
    if (chartType === 'line') chartInstructionsTextLine.value = msg
    if (chartType === 'pie') chartInstructionsTextPie.value = msg
    if (chartType === 'bar') chartInstructionsTextBar.value = msg
    return
  }

  const parsed = parseRange(dataRangeInput.value || dataRange.value, dataMatrix.value.length, dataMatrix.value[0]?.length || 0)
  if (!parsed) {
    rangeError.value = t('invalidRange')
    return
  }
  if ((parsed as any).clamped) {
    const safeRange = formatRange(parsed)
    dataRangeInput.value = safeRange
    rangeError.value = t('rangeClamped', { range: safeRange })
  }

  const { labels, values } = extractSeries(dataMatrix.value, targetColumn, parsed)
  if (!labels.length || !values.length) {
    const msg = t('noValidData')
    if (chartType === 'line') chartInstructionsTextLine.value = msg
    if (chartType === 'pie') chartInstructionsTextPie.value = msg
    if (chartType === 'bar') chartInstructionsTextBar.value = msg
    return
  }

  if (chartType === 'line' && lineCanvasRef.value) drawLine(lineCanvasRef.value, labels, values)
  if (chartType === 'pie' && pieCanvasRef.value) drawPie(pieCanvasRef.value, labels, values)
  if (chartType === 'bar' && barCanvasRef.value) drawBar(barCanvasRef.value, labels, values)

  const successMsg = chartType === 'line'
    ? t('chartGeneratedLine', { column: targetColumn })
    : chartType === 'pie'
      ? t('chartGeneratedPie', { column: targetColumn })
      : t('chartGeneratedBar', { column: targetColumn })
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
    alert(t('recommendFailed'))
  }
}

async function onGenerateReportClick() {
  if (!savedFile.value) {
    alert(t('reportNeedsFile'))
    return
  }

  reportGenerating.value = true
  reportGenerated.value = true
  reportMarkdown.value = '生成中...'
  const file = savedFile.value
  const sections: string[] = []
  let previewFinal = ''
  const hasOptional = includeFinancialRatios.value || includeProfitability.value || includeCashFlow.value || includeBudgetActual.value || includeRfm.value || includeClv.value

  try {
    // 基础：财务报表分析（必调）
    try {
      const res = await analyzeFinancial(file)
      if ((res as any)?.excelDataPreview) previewFinal = (res as any).excelDataPreview
      sections.push(renderSectionMd(t('reportSectionFinancial'), (res as any)?.financialAnalysis || (res as any)?.analysis, undefined))
    } catch (err) {
      sections.push(renderErrorMd(t('reportSectionFinancial'), err))
    }

    // 选填：财务比率、盈利能力、现金流、预算对比、RFM、CLV
    if (hasOptional) {
      if (includeFinancialRatios.value) {
        try {
          const res = await analyzeFinancialRatios(file)
          if ((res as any)?.excelDataPreview) previewFinal = (res as any).excelDataPreview
          sections.push(renderSectionMd(t('reportSectionFinancialRatios'), (res as any)?.financialRatios, undefined))
        } catch (err) {
          sections.push(renderErrorMd(t('reportSectionFinancialRatios'), err))
        }
      }

      if (includeProfitability.value) {
        try {
          const res = await analyzeProfitability(file)
          if ((res as any)?.excelDataPreview) previewFinal = (res as any).excelDataPreview
          sections.push(renderSectionMd(t('reportSectionProfitability'), (res as any)?.profitabilityAnalysis, undefined))
        } catch (err) {
          sections.push(renderErrorMd(t('reportSectionProfitability'), err))
        }
      }

      if (includeCashFlow.value) {
        try {
          const res = await analyzeCashFlow(file)
          if ((res as any)?.excelDataPreview) previewFinal = (res as any).excelDataPreview
          sections.push(renderSectionMd(t('reportSectionCashFlow'), (res as any)?.cashFlowAnalysis, undefined))
        } catch (err) {
          sections.push(renderErrorMd(t('reportSectionCashFlow'), err))
        }
      }

      if (includeBudgetActual.value) {
        try {
          const res = await analyzeBudgetActual(file)
          if ((res as any)?.excelDataPreview) previewFinal = (res as any).excelDataPreview
          sections.push(renderSectionMd(t('reportSectionBudgetActual'), (res as any)?.budgetActualAnalysis, undefined))
        } catch (err) {
          sections.push(renderErrorMd(t('reportSectionBudgetActual'), err))
        }
      }

      if (includeRfm.value) {
        try {
          const res = await analyzeRfm(file)
          if ((res as any)?.excelDataPreview) previewFinal = (res as any).excelDataPreview
          sections.push(renderSectionMd(t('reportSectionRfm'), (res as any)?.rfmAnalysis, undefined))
        } catch (err) {
          sections.push(renderErrorMd(t('reportSectionRfm'), err))
        }
      }

      if (includeClv.value) {
        try {
          const res = await analyzeClv(file)
          if ((res as any)?.excelDataPreview) previewFinal = (res as any).excelDataPreview
          sections.push(renderSectionMd(t('reportSectionClv'), (res as any)?.clvAnalysis, undefined))
        } catch (err) {
          sections.push(renderErrorMd(t('reportSectionClv'), err))
        }
      }
    }

    if (previewFinal) {
      sections.push(`**${t('reportPreviewTitle')}**\n\n\u0060\u0060\u0060\n${previewFinal}\n\u0060\u0060\u0060`)
    }

    reportMarkdown.value = sections.join('\n\n') || t('reportNoContent')
    reportGenerated.value = true

    // 等待 DOM 更新然后测量实际高度
    await nextTick()
    const el = document.querySelector('.auto-report-card .report-content') as HTMLElement | null
    if (el) {
      reportHeight.value = el.scrollHeight
    } else {
      reportHeight.value = 800
    }
  } finally {
    reportGenerating.value = false
  }
}

async function onDownloadReport() {
  if (!reportGenerated.value || !reportMarkdown.value) {
    alert(t('reportNoContent'))
    return
  }
  const el = document.querySelector('.auto-report-card .report-content') as HTMLElement | null
  if (!el) {
    alert(t('reportNoContent'))
    return
  }
  try {
    const card = document.querySelector('.auto-report-card') as HTMLElement | null
    const bgColor = '#19202C'
    const headings = Array.from(el.querySelectorAll('h3')) as HTMLElement[]
    const sections: HTMLElement[] = []

    if (!headings.length) {
      sections.push(el.cloneNode(true) as HTMLElement)
    } else {
      const headingSet = new Set(headings)
      headings.forEach((h, idx) => {
        const section = document.createElement('div')
        section.style.padding = '12px'
        section.style.background = bgColor
        section.style.width = `${el.clientWidth || el.offsetWidth || 800}px`
        let node: Node | null = h
        while (node) {
          section.appendChild(node.cloneNode(true))
          node = node.nextSibling
          if (node && node.nodeType === 1 && headingSet.has(node as HTMLElement)) break
        }
        sections.push(section)
      })
    }

    const staging = document.createElement('div')
    staging.style.position = 'fixed'
    staging.style.left = '-99999px'
    staging.style.top = '0'
    staging.style.zIndex = '-1'
    staging.style.background = bgColor
    document.body.appendChild(staging)

    const pdf = new jsPDF('p', 'pt', 'a4')
    const pageWidth = pdf.internal.pageSize.getWidth()
    const pageHeight = pdf.internal.pageSize.getHeight()
    const margin = 20
    pdf.setFillColor(25, 32, 44)
    pdf.rect(0, 0, pageWidth, pageHeight, 'F')
    let cursorY = margin

    for (const section of sections) {
      staging.appendChild(section)
      const canvas = await html2canvas(section, { scale: 2, useCORS: true, backgroundColor: bgColor })
      const imgData = canvas.toDataURL('image/png')
      const imgWidth = pageWidth - margin * 2
      const imgHeight = (canvas.height * imgWidth) / canvas.width

      if (cursorY + imgHeight > pageHeight - margin) {
        pdf.addPage()
        pdf.setFillColor(25, 32, 44)
        pdf.rect(0, 0, pageWidth, pageHeight, 'F')
        cursorY = margin
      }

      pdf.addImage(imgData, 'PNG', margin, cursorY, imgWidth, imgHeight)
      cursorY += imgHeight + 10

      staging.removeChild(section)
    }

    document.body.removeChild(staging)

    const blob = pdf.output('blob')
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = 'auto-report.pdf'
    link.click()
    URL.revokeObjectURL(url)
  } catch (err) {
    console.error('download report failed', err)
    alert(t('downloadFailed') || '下载失败')
  }
}

function onRfmClick() {
  alert(t('rfmClicked'))
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

  reportMarkdown.value = '```json\n' + JSON.stringify(example, null, 2) + '\n```'
  reportGenerated.value = true
  // 等待 DOM 更新然后测量高度
  nextTick().then(() => {
    const el = document.querySelector('.auto-report-card .report-content') as HTMLElement | null
    if (el) reportHeight.value = el.scrollHeight
    else reportHeight.value = 800
  })
}

function onClvClick() {
  alert(t('clvClicked'))
}

function onFinanceClick() {
  alert(t('financeClicked'))
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
.download-report-btn { margin-left: auto; white-space: nowrap; }

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
  min-width: 0;
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
.auto-report-card .report-placeholder { display:flex; align-items:center; justify-content:center; height:100%; color: rgba(230,238,248,0.9); }
.auto-report-card .report-content { padding:12px; color: rgba(230,238,248,0.9); }
.auto-report-card .report-content :deep(.markdown-body) {
  color: rgba(230,238,248,0.9);
  background: transparent;
}
.auto-report-card .report-content :deep(.markdown-body p),
.auto-report-card .report-content :deep(.markdown-body li),
.auto-report-card .report-content :deep(.markdown-body span),
.auto-report-card .report-content :deep(.markdown-body strong),
.auto-report-card .report-content :deep(.markdown-body em) {
  color: rgba(230,238,248,0.9);
}
.auto-report-card .report-content :deep(.markdown-body h1),
.auto-report-card .report-content :deep(.markdown-body h2),
.auto-report-card .report-content :deep(.markdown-body h3),
.auto-report-card .report-content :deep(.markdown-body h4),
.auto-report-card .report-content :deep(.markdown-body h5),
.auto-report-card .report-content :deep(.markdown-body h6) {
  color: rgba(230,238,248,0.95);
}
.auto-report-card .report-content :deep(.markdown-body a) {
  color: #9ad4ff;
}
.auto-report-card .report-content :deep(.elx-xmarkdown-container),
.auto-report-card .report-content :deep(.elx-xmarkdown-provider),
.auto-report-card .report-content :deep(.elx-xmarkdown-container *),
.auto-report-card .report-content :deep(.elx-xmarkdown-provider *) {
  color: rgba(230,238,248,0.9);
}
.auto-report-card .report-content :deep(pre),
.auto-report-card .report-content :deep(code),
.auto-report-card .report-content :deep(.elx-highlight-code-wrapper),
.auto-report-card .report-content :deep(.elx-highlight-code-wrapper *),
.auto-report-card .report-content :deep(.shiki),
.auto-report-card .report-content :deep(.shiki *) {
  color: rgba(230,238,248,0.9);
  background-color: rgba(255,255,255,0.02) !important;
}
.auto-report-card .report-content :deep(.pre-md),
.auto-report-card .report-content :deep(.markdown-elxLanguage-header-div),
.auto-report-card .report-content :deep(.markdown-elxLanguage-header-span),
.auto-report-card .report-content :deep(.markdown-elxLanguage-header-space),
.auto-report-card .report-content :deep(.el-button.shiki-header-button),
.auto-report-card .report-content :deep(.el-button.shiki-header-button-expand),
.auto-report-card .report-content :deep(.el-button.shiki-header-button-text),
.auto-report-card .report-content :deep(.el-scrollbar__wrap),
.auto-report-card .report-content :deep(.el-scrollbar__view),
.auto-report-card .report-content :deep(.code-lines) {
  background: rgba(255,255,255,0.02) !important;
  color: rgba(230,238,248,0.9) !important;
  border-color: rgba(255,255,255,0.06) !important;
}
.auto-report-card .report-content :deep(.code-lines .line-content),
.auto-report-card .report-content :deep(.code-lines .line),
.auto-report-card .report-content :deep(.code-lines .line span) {
  color: rgba(230,238,248,0.9) !important;
}
.auto-report-card .report-content :deep(.el-scrollbar__bar) {
  background: transparent !important;
}
.auto-report-card .report-content :deep(.el-scrollbar__thumb) {
  background: rgba(255,255,255,0.25) !important;
  border-radius: 999px !important;
}
.auto-report-card .report-content :deep(.markdown-elxLanguage-header-div) {
  box-shadow: none !important;
  background: transparent !important;
}
.auto-report-card .report-content :deep(.shiki-header-button),
.auto-report-card .report-content :deep(.markdown-elxLanguage-header-toggle),
.auto-report-card .report-content :deep(.markdown-elxLanguage-header-button) {
  background: transparent !important;
  border: none !important;
  box-shadow: none !important;
}
.auto-report-card .report-content :deep(.shiki .line),
.auto-report-card .report-content :deep(.shiki .line span),
.auto-report-card .report-content :deep(.shiki .line-content) {
  color: rgba(230,238,248,0.9) !important;
}
.auto-report-card .report-content :deep(table),
.auto-report-card .report-content :deep(thead),
.auto-report-card .report-content :deep(tbody),
.auto-report-card .report-content :deep(tr),
.auto-report-card .report-content :deep(th),
.auto-report-card .report-content :deep(td) {
  background: transparent !important;
  color: rgba(230,238,248,0.9) !important;
  border-color: rgba(255,255,255,0.06) !important;
}
.auto-report-card .report-content :deep(th) {
  font-weight: 700;
}
.auto-report-card.generated { /* extra visual emphasis when generated */ box-shadow: 0 2px 10px rgba(0,0,0,0.3); }
.report-actions { display:flex; flex-direction:column; gap:8px; margin:8px 0; }
.report-buttons { display:flex; gap:8px; flex-wrap:wrap; }
.report-options { display:flex; gap:16px; flex-wrap:nowrap; align-items:center; font-size:12px; color: rgba(230,238,248,0.8); }
.report-options label { display:flex; gap:6px; align-items:center; cursor:pointer; white-space:nowrap; }
.report-options .options-label { font-weight:600; color: rgba(230,238,248,0.9); }
.report-block { border:1px solid rgba(255,255,255,0.06); border-radius:6px; padding:12px; margin-bottom:12px; background: rgba(255,255,255,0.02); }
.report-block h3 { margin:0 0 6px 0; font-size:16px; }
.report-block p { margin:0 0 6px 0; line-height:1.5; }
.report-block.error { border-color: rgba(255,120,120,0.4); background: rgba(255,120,120,0.05); }
.report-preview { margin-top:8px; background: rgba(255,255,255,0.03); border:1px solid rgba(255,255,255,0.05); border-radius:4px; }
.report-preview-label { padding:6px 8px; border-bottom:1px solid rgba(255,255,255,0.06); font-size:12px; color: rgba(230,238,248,0.75); }
.report-preview pre { margin:0; padding:8px; color: rgba(230,238,248,0.85); white-space:pre-wrap; word-break:break-all; }

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