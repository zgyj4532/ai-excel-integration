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
        <p class="text-muted" style="margin-top:8px">{{ $t('uploaderHint') || '拖拽文件到此处或点击下方按钮上传' }}</p>

        <input
          type="file"
          id="fileInput"
          ref="fileInput"
          class="d-none"
          accept=".xlsx,.xls,.csv"
          @change="onChange"
        />

        <button class="btn btn-primary" id="browseFileBtn" @click="openBrowser">
          <i class="bi bi-folder2-open me-2"></i>{{ $t('chooseFile') || '选择文件' }}
        </button>

        <div v-if="selectedName" style="margin-top:8px" class="muted">{{ $t('currentFile') }} {{ selectedName }}</div>
      </div>

      <div v-if="progress>0" style="margin-top:8px">
        <div class="muted">{{ $t('parsingProgress', { pct: Math.round(progress*100) }) }}</div>
        <div style="background:rgba(255,255,255,0.04); height:6px; border-radius:4px; overflow:hidden; margin-top:6px">
          <div :style="{ width: (progress*100)+'%', background: 'linear-gradient(90deg,#7c3aed,#38bdf8)', height:'6px' }"></div>
        </div>
      </div>

      <div style="margin-top:8px; display:flex; gap:8px; align-items:center; flex-wrap:wrap">
        <button @click="downloadSample">{{ $t('uploaderDownloadSample') }}</button>
        <button @click="downloadSampleXlsx">下载示例 XLSX</button>
        <div class="muted">{{ $t('uploaderSupport', { size: 5 }) }}</div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'
const { t } = useI18n()
import * as XLSX from 'xlsx'
const emit = defineEmits(['fileLoaded'])

const progress = ref(0)
const selectedName = ref('')
const fileInput = ref<HTMLInputElement | null>(null)
const isDragOver = ref(false)

function openBrowser(){
  fileInput.value && fileInput.value.click()
}

function onDragEnter(e: DragEvent){
  isDragOver.value = true
  if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy'
}

function onDragOver(e: DragEvent){
  if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy'
}

function onDragLeave(){
  isDragOver.value = false
}

function parseCSV(text: string){
  const cleaned = text.replace(/^\uFEFF/, '')
  return cleaned.split(/\r?\n/).filter(Boolean).map(line => line.split(/,|\t/))
}

function decodeBufferToText(buf: ArrayBuffer){
  // quick BOM detection
  const u8 = new Uint8Array(buf)
  if (u8.length >= 2) {
    // UTF-16LE BOM
    if (u8[0] === 0xFF && u8[1] === 0xFE) {
      try { return new TextDecoder('utf-16le').decode(buf).replace(/^\uFEFF/, '') } catch (e) { /* ignore */ }
    }
    // UTF-16BE BOM
    if (u8[0] === 0xFE && u8[1] === 0xFF) {
      try { return new TextDecoder('utf-16be').decode(buf).replace(/^\uFEFF/, '') } catch (e) { /* ignore */ }
    }
  }

  const decoders = ['utf-8', 'gb18030', 'gbk', 'gb2312', 'big5', 'utf-16le', 'utf-16be']
  let bestText = ''
  let bestScore = -Infinity

  for (const enc of decoders){
    try {
      const dec = new TextDecoder(enc as any, { fatal: false })
      const txt = dec.decode(buf)
      const replacements = (txt.match(/\uFFFD/g) || []).length
      const cjk = (txt.match(/[\u4E00-\u9FFF]/g) || []).length
      const asciiLike = (txt.match(/[\x20-\x7E]/g) || []).length
      // score: prefer fewer replacements, more CJK, reasonable ascii
      const score = cjk * 5 + asciiLike * 0.1 - replacements * 1000
      if (score > bestScore) {
        bestScore = score
        bestText = txt
      }
      if (replacements === 0 && cjk > 0) break
    } catch (e) {
      /* ignore unsupported encoding */
    }
  }

  if (!bestText) {
    try { bestText = new TextDecoder('utf-8').decode(buf) } catch (e) { bestText = '' }
  }

  return bestText.replace(/^\uFEFF/, '')
}

function parseXLSX(buffer: ArrayBuffer){
  const wb = XLSX.read(buffer, { type: 'array' })
  const first = wb.SheetNames[0]
  const sheet = wb.Sheets[first]
  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[]
  return aoa.map(r => r.map((c:any)=> c==null ? '' : String(c)))
}

function handleFile(file: File){
  if (file.size > 5 * 1024 * 1024) {
    alert(t('fileTooLarge', { size: 5 }))
    return
  }

  progress.value = 0.05

  if (/\.xls|\.xlsx$/i.test(file.name)){
    const reader = new FileReader()
    reader.onload = () => {
      try{
        const buf = reader.result as ArrayBuffer
        progress.value = 0.6
        const data = parseXLSX(buf)
        progress.value = 1
        emit('fileLoaded', { name: file.name, data, file })
        selectedName.value = file.name
        setTimeout(()=> progress.value = 0, 500)
      }catch(e){ alert(t('parseFailed')); progress.value = 0 }
    }
    reader.onprogress = (e)=>{ if (e.lengthComputable) progress.value = e.loaded / e.total }
    reader.readAsArrayBuffer(file)
    return
  }

  // fallback CSV (decode with charset detection to avoid garbled Chinese)
  const reader = new FileReader()
  reader.onload = () => {
    const buffer = reader.result as ArrayBuffer
    const text = decodeBufferToText(buffer)
    const data = parseCSV(text)
    progress.value = 1
    emit('fileLoaded', { name: file.name, data, file })
    selectedName.value = file.name
    setTimeout(()=> progress.value = 0, 500)
  }
  reader.onprogress = (e)=>{ if (e.lengthComputable) progress.value = e.loaded / e.total }
  reader.readAsArrayBuffer(file)
}

function onChange(e: Event){
  const input = e.target as HTMLInputElement
  const file = input.files && input.files[0]
  if (!file) return
  handleFile(file)
}

function onDrop(e: DragEvent){
  isDragOver.value = false
  const file = e.dataTransfer && e.dataTransfer.files[0]
  if (!file) return
  handleFile(file)
}

function downloadSample(){
  const csv = [
    '年份,收入,成本,营业费用,净利润',
    '2018,300000,180000,90000,30000',
    '2019,320000,190000,95000,35000',
    '2020,350000,210000,100000,40000',
    '2021,380000,230000,110000,40000',
    '2022,400000,240000,120000,40000',
    '2023,430000,260000,130000,40000',
    '2024,450000,270000,135000,45000',
    '2025,480000,290000,145000,45000',
    '2026,500000,300000,150000,50000'
  ].join('\n')
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = 'sample.csv'
  document.body.appendChild(a)
  a.click()
  a.remove()
  URL.revokeObjectURL(url)
}

function downloadSampleXlsx(){
  const data = [
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
  const ws = XLSX.utils.aoa_to_sheet(data)
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = 'sample.xlsx'
  document.body.appendChild(a)
  a.click()
  a.remove()
  URL.revokeObjectURL(url)
}
</script>
