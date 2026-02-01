<script setup lang="ts">
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'
import * as XLSX from 'xlsx'

type SnapshotOps = {
  createWorkbook?: (snapshot?: any) => any
}

const props = withDefaults(defineProps<{
  snapshotOps: SnapshotOps | null
}>(), {
  snapshotOps: null,
})

const emit = defineEmits<{ (e: 'fileLoaded', payload: { name: string; data: string[][]; file: File }): void }>()

const { t } = useI18n()
const uploading = ref(false)
const message = ref('')
const fileInput = ref<HTMLInputElement | null>(null)
const isDragOver = ref(false)
const progress = ref(0)
const selectedName = ref('')

const makeId = (prefix: string) => `${prefix}-${Math.random().toString(36).slice(2, 8)}`

const sheetToWorksheetData = (sheet: XLSX.WorkSheet, name: string, sheetId: string) => {
  const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true }) as any[][]
  const rowCount = rows.length || 1
  const columnCount = rows.reduce((m, r) => Math.max(m, r.length), 0) || 1
  const cellData: Record<string, Record<string, { v: any }>> = {}

  rows.forEach((r, rIdx) => {
    r.forEach((value, cIdx) => {
      if (value === undefined || value === null || value === '') return
      const rowKey = String(rIdx)
      if (!cellData[rowKey]) cellData[rowKey] = {}
      cellData[rowKey][String(cIdx)] = { v: value }
    })
  })

  return {
    id: sheetId,
    name,
    rowCount,
    columnCount,
    cellData,
  }
}

const workbookToSnapshot = (workbook: XLSX.WorkBook, fileName: string) => {
  const sheetOrder: string[] = []
  const sheets: Record<string, any> = {}

  workbook.SheetNames.forEach((name) => {
    const sheet = workbook.Sheets[name]
    if (!sheet) return
    const sheetId = makeId('sheet')
    sheetOrder.push(sheetId)
    sheets[sheetId] = sheetToWorksheetData(sheet, name, sheetId)
  })

  return {
    id: makeId('workbook'),
    name: fileName,
    sheetOrder,
    sheets,
  }
}

const parsePreview = (workbook: XLSX.WorkBook) => {
  const first = workbook.SheetNames[0]
  const sheet = first ? workbook.Sheets[first] : null
  if (!sheet) return [] as string[][]
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true }) as any[]
  return aoa.map((r) => r.map((c: any) => (c == null ? '' : String(c))))
}

const openBrowser = () => {
  fileInput.value?.click()
}

const processFile = async (file: File) => {
  const ops = props.snapshotOps
  const hasCreate = ops && typeof ops.createWorkbook === 'function'

  uploading.value = true
  message.value = ''
  progress.value = 0

  try {
    progress.value = 0.3
    const buffer = await file.arrayBuffer()
    progress.value = 0.6
    const workbook = XLSX.read(buffer, { type: 'array' })
    const snapshot = workbookToSnapshot(workbook, file.name)
    const preview = parsePreview(workbook)
    progress.value = 0.9

    if (!hasCreate) {
      message.value = t('uploadSuccess')
    } else {
      try {
        ops?.createWorkbook?.(snapshot)
        message.value = t('uploadSuccess')
      } catch (error) {
        const errMsg = (error as Error)?.message || t('unknownError')
        const fallback = errMsg.toLowerCase().includes('disposed') ? t('uploadSuccess') : errMsg
        message.value = t('uploadFailedWithReason', { reason: fallback })
        return
      }
    }

    selectedName.value = file.name
    emit('fileLoaded', { name: file.name, data: preview, file })
    progress.value = 1
    setTimeout(() => { progress.value = 0 }, 500)
  } catch (error) {
    const reason = error instanceof Error ? error.message : t('unknownError')
    message.value = t('uploadFailedWithReason', { reason })
    progress.value = 0
  } finally {
    uploading.value = false
  }
}

const onChange = (event: Event) => {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (!file) return
  processFile(file)
  target.value = ''
}

const onDragEnter = () => {
  isDragOver.value = true
}

const onDragOver = () => {
  isDragOver.value = true
}

const onDragLeave = () => {
  isDragOver.value = false
}

const onDrop = (event: DragEvent) => {
  isDragOver.value = false
  const files = event.dataTransfer?.files
  if (!files || files.length === 0) return
  const file = files[0]
  if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.type === 'application/vnd.ms-excel' ||
      file.type === 'text/csv' ||
      file.name.endsWith('.xlsx') ||
      file.name.endsWith('.xls') ||
      file.name.endsWith('.csv')) {
    processFile(file)
  } else {
    message.value = t('unsupportedFormat')
  }
}

const downloadSample = () => {
  const sample = [
    ['年份','收入','成本','营业费用','净利润'],
    ['2018',300000,180000,90000,30000],
    ['2019',320000,190000,95000,35000],
    ['2020',350000,210000,100000,40000],
    ['2021',380000,230000,110000,40000],
    ['2022',400000,240000,120000,40000],
    ['2023',430000,260000,130000,40000],
    ['2024',450000,270000,135000,45000],
    ['2025',480000,290000,145000,45000],
    ['2026',500000,300000,150000,50000]
  ]
  const ws = XLSX.utils.aoa_to_sheet(sample)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
  XLSX.write(wb, { bookType: 'csv', type: 'string' })
  const blob = new Blob([XLSX.write(wb, { bookType: 'csv', type: 'string' })], { type: 'text/csv;charset=utf-8;' })
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob)
  link.download = 'sample.csv'
  link.click()
  URL.revokeObjectURL(link.href)
}

const downloadSampleXlsx = () => {
  const sample = [
    ['年份','收入','成本','营业费用','净利润'],
    ['2018',300000,180000,90000,30000],
    ['2019',320000,190000,95000,35000],
    ['2020',350000,210000,100000,40000],
    ['2021',380000,230000,110000,40000],
    ['2022',400000,240000,120000,40000],
    ['2023',430000,260000,130000,40000],
    ['2024',450000,270000,135000,45000],
    ['2025',480000,290000,145000,45000],
    ['2026',500000,300000,150000,50000]
  ]
  const ws = XLSX.utils.aoa_to_sheet(sample)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
  XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
  const blob = new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob)
  link.download = 'sample.xlsx'
  link.click()
  URL.revokeObjectURL(link.href)
}
</script>

<template>
  <div>
    <label class="muted">{{ $t('uploaderLabel') }}</label>

    <div class="card" style="padding:12px">
      <div
        :class="['file-upload-area', { dragover: isDragOver }]"
        id="fileDropArea"
        @drop.prevent="onDrop"
        @dragenter.prevent="onDragEnter"
        @dragover.prevent="onDragOver"
        @dragleave.prevent="onDragLeave"
        style="padding:18px; text-align:center; cursor:copy"
      >
        <svg class="cloud-icon" width="48" height="48" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M4.406 3.342A4.002 4.002 0 0 1 11.5 4.5a3.5 3.5 0 0 1 .5 6.992v.008H4.5a3 3 0 0 1-.094-5.666 4.002 4.002 0 0 1 0-2.492z" fill="currentColor" opacity="0.12"/>
          <path d="M8 5.5v4M6 7.5l2-2 2 2" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <p class="text-muted" style="margin-top:8px">{{ $t('uploaderHint') }}</p>

        <input
          type="file"
          id="fileInput"
          ref="fileInput"
          class="d-none"
          accept=".xlsx,.xls,.csv"
          @change="onChange"
        />

        <button class="btn btn-primary" id="browseFileBtn" @click="openBrowser" :disabled="uploading">
          <span style="margin-right:6px">📁</span>{{ $t('chooseFile') }}
        </button>

        <div v-if="selectedName" style="margin-top:8px" class="muted">{{ $t('currentFile') }} {{ selectedName }}</div>
      </div>

      <div v-if="progress>0" style="margin-top:8px">
        <div class="muted">{{ $t('parsingProgress', { pct: Math.round(progress*100) }) }}</div>
        <div style="background:rgba(255,255,255,0.04); height:6px; border-radius:4px; overflow:hidden; margin-top:6px">
          <div :style="{ width: (progress*100)+'%', background: 'linear-gradient(90deg,#7c3aed,#38bdf8)', height:'6px' }"></div>
        </div>
      </div>

      <div v-if="message" style="margin-top:8px; padding:8px; border-radius:4px; background:rgba(255,255,255,0.04); color:#9ca3af; font-size:12px">
        {{ message }}
      </div>

      <div style="margin-top:8px; display:flex; gap:8px; align-items:center; flex-wrap:wrap">
        <button class="btn-link" @click="downloadSample">{{ $t('uploaderDownloadSample') }}</button>
        <button class="btn-link" @click="downloadSampleXlsx">{{ $t('uploaderDownloadSampleXlsx') }}</button>
        <div class="muted">{{ $t('uploaderSupport', { size: 5 }) }}</div>
      </div>
    </div>
  </div>
</template>

<style scoped>
.muted {
  color: #9ca3af;
  font-size: 13px;
}

.text-muted {
  color: #9ca3af;
  font-size: 14px;
  margin: 0;
}

.card {
  background: transparent;
  border: 1px solid #374151;
  border-radius: 6px;
}

.file-upload-area {
  border: 2px dashed #4b5563;
  border-radius: 6px;
  background: rgba(255, 255, 255, 0.01);
  transition: all 0.2s ease;
}

.file-upload-area:hover {
  border-color: #6b7280;
  background: rgba(255, 255, 255, 0.02);
}

.file-upload-area.dragover {
  border-color: #38bdf8;
  background: rgba(56, 189, 248, 0.05);
}

.cloud-icon {
  color: #6b7280;
  transition: color 0.2s ease;
}

.file-upload-area.dragover .cloud-icon {
  color: #38bdf8;
}

.d-none {
  display: none;
}

.btn {
  padding: 8px 16px;
  background: #189079;
  color: #fff;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  font-size: 14px;
  transition: background 0.2s ease;
}

.btn:hover:not(:disabled) {
  background: #15754f;
}

.btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.btn-primary {
  background: #189079;
  color: #fff;
}

.btn-link {
  padding: 8px 16px;
  background: #189079;
  color: #fff;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  font-size: 14px;
  transition: background 0.2s ease;
}

.btn-link:hover:not(:disabled) {
  background: #15754f;
}

.btn-link:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>
