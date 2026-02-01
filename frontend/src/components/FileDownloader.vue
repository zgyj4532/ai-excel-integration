<script setup lang="ts">
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'
import * as XLSX from 'xlsx'

type SnapshotOps = {
  getActiveWorkbook?: () => any
}

const props = defineProps<{
  snapshotOps: SnapshotOps | null
  isOpen: boolean
}>()

const emit = defineEmits<{ (e: 'close'): void }>()

const downloading = ref(false)
const message = ref('')
const format = ref<'xlsx' | 'csv' | 'xls'>('xlsx')
const { t } = useI18n()

const downloadBlob = (file: Blob, filename: string) => {
  const url = URL.createObjectURL(file)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(url)
}

const buildSheetFromSnapshot = (snapshot: any) => {
  const sheetOrder = snapshot?.sheetOrder || []
  const sheets = snapshot?.sheets || {}
  const firstSheetId = sheetOrder[0]
  if (!firstSheetId || !sheets[firstSheetId]) throw new Error('未找到工作表数据')

  const sheetData = sheets[firstSheetId]
  const rows = sheetData.rowCount || 1
  const cols = sheetData.columnCount || 1
  const table: any[][] = Array.from({ length: rows }, () => Array(cols).fill(undefined))
  const cellData = sheetData.cellData || {}
  Object.keys(cellData).forEach((rKey) => {
    const row = Number(rKey)
    const rowCells = cellData[rKey]
    if (!rowCells) return
    if (!table[row]) return
    Object.keys(rowCells).forEach((cKey) => {
      const col = Number(cKey)
      if (table[row] && rowCells[cKey]) {
        table[row][col] = rowCells[cKey]?.v
      }
    })
  })
  return XLSX.utils.aoa_to_sheet(table)
}

const handleDownload = async () => {
  const ops = props.snapshotOps
  if (!ops || !ops.getActiveWorkbook) {
    message.value = t('exportNotReady')
    return
  }

  downloading.value = true
  message.value = ''

  try {
    const workbook = ops.getActiveWorkbook()
    if (!workbook) {
      throw new Error(t('exportNoWorkbook'))
    }

    const getSnapshot = typeof workbook.getSnapshot === 'function' ? workbook.getSnapshot.bind(workbook) : null
    if (!getSnapshot) {
      throw new Error(t('exportNoSnapshot'))
    }

    const snapshot = getSnapshot()
    if (!snapshot) {
      throw new Error(t('exportNoSnapshot'))
    }

    const wb = XLSX.utils.book_new()
    const ws = buildSheetFromSnapshot(snapshot)
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')

    const bookType = format.value
    const out = XLSX.write(wb, { bookType, type: bookType === 'csv' ? 'string' : 'array' })
    const blob = bookType === 'csv'
      ? new Blob([out as string], { type: 'text/csv' })
      : new Blob([out as ArrayBuffer], { type: 'application/vnd.ms-excel' })

    downloadBlob(blob, `univer.${bookType}`)
    message.value = t('exportSuccess')
    setTimeout(() => emit('close'), 800)
  } catch (error) {
    const reason = error instanceof Error ? error.message : t('unknownError')
    message.value = t('exportFileFailed', { reason })
  } finally {
    downloading.value = false
  }
}
</script>

<template>
  <Teleport to="body" v-if="isOpen">
    <div class="modal-overlay" @click="emit('close')">
      <div class="modal" @click.stop>
        <div class="modal-header">
          <h2>{{ t('exportModalTitle') }}</h2>
          <button class="close-btn" @click="emit('close')" :aria-label="t('exportClose')">×</button>
        </div>
        <div class="modal-body">
          <label class="select-row">
            <span>{{ t('exportFormat') }}</span>
            <select v-model="format" class="select">
              <option value="xlsx">{{ t('exportXlsx') }}</option>
              <option value="csv">{{ t('exportCsv') }}</option>
              <option value="xls">{{ t('exportXls') }}</option>
            </select>
          </label>
        </div>
        <div class="modal-footer">
          <button class="btn-cancel" @click="emit('close')">{{ t('exportCancel') }}</button>
          <button class="btn-download" :disabled="downloading" @click="handleDownload">
            {{ downloading ? t('exportDownloading') : t('exportDownload') }}
          </button>
        </div>
        <p v-if="message" class="msg">{{ message }}</p>
      </div>
    </div>
  </Teleport>
</template>

<style scoped>
.modal-overlay {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(0, 0, 0, 0.5);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 1000;
}

.modal {
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.15);
  max-width: 400px;
  width: 90%;
  padding: 24px;
}

.modal-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 16px;
}

.modal-header h2 {
  margin: 0;
  font-size: 18px;
  color: #1f2933;
}

.close-btn {
  background: none;
  border: none;
  font-size: 28px;
  cursor: pointer;
  color: #6b7280;
  padding: 0;
  width: 32px;
  height: 32px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.close-btn:hover {
  color: #1f2933;
}

.modal-body {
  margin-bottom: 20px;
}

.select-row {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #374151;
}

.select-row span {
  min-width: 70px;
}

.select {
  flex: 1;
  padding: 8px 10px;
  border-radius: 6px;
  border: 1px solid #d1d5db;
  background: #fff;
  font-size: 13px;
  cursor: pointer;
}

.modal-footer {
  display: flex;
  gap: 8px;
  justify-content: flex-end;
}

.btn-cancel {
  padding: 10px 16px;
  background: #f3f4f6;
  color: #374151;
  border: 1px solid #d1d5db;
  border-radius: 6px;
  cursor: pointer;
  font-size: 14px;
  transition: background 0.2s ease;
}

.btn-cancel:hover {
  background: #e5e7eb;
}

.btn-download {
  padding: 10px 16px;
  background: #189079;
  color: #fff;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  font-size: 14px;
  transition: background 0.2s ease;
}

.btn-download:hover:not(:disabled) {
  background: #15754f;
}

.btn-download:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.msg {
  margin: 12px 0 0;
  font-size: 12px;
  color: #6b7280;
  text-align: center;
}
</style>
