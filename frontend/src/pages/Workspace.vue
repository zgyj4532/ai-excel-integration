<template>
  <div style="padding:16px">
    <!-- workspace grid: left column contains topbar(row1) + preview(row2); right column is sidebar -->
    <div class="workspace-grid">
      <div class="topbar">
            <div class="muted">{{ $t('currentFile') }} {{ fileName || $t('noFile') }}</div>
            <div style="margin-left:auto; display:flex; gap:8px; align-items:center;">
              <button class="new-chat-btn" @click="newChat" :title="t('newChat')" :aria-label="t('newChat')">
                <span class="new-chat-circle" aria-hidden="true">
                  <svg class="new-chat-plus" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" role="img"
                    aria-hidden="true">
                    <line x1="12" y1="5" x2="12" y2="19" stroke="#179078" stroke-width="2" stroke-linecap="round" />
                    <line x1="5" y1="12" x2="19" y2="12" stroke="#179078" stroke-width="2" stroke-linecap="round" />
                  </svg>
                </span>
                <span class="new-chat-label">{{ t('newChat') }}</span>
              </button>
              
            </div>
            <!-- Error banner with retry -->
            <div v-if="lastError" style="margin-top:8px; display:flex; gap:8px; align-items:center;">
              <div style="color:#ffcccc; background:rgba(255,0,0,0.06); padding:8px; border-radius:6px; flex:1">{{ lastError }}</div>
              <button class="btn" @click="retryLastCommand" v-if="failedCommand">{{ t('retry') || '重试' }}</button>
            </div>
          </div>
      <section class="preview card" style="flex:1; min-width:0; height:100%; overflow:auto;">
        <div
          style="display:flex; align-items:center; gap:8px; padding:12px 16px; border-bottom:1px solid rgba(255,255,255,0.03)">
          <div style="display:flex; gap:8px">
            <button :class="['tab-btn', activeTab === 'overview' ? 'active' : '']" @click="activeTab = 'overview'">{{
              $t('dataPreview') }}</button>
            <button :class="['tab-btn', activeTab === 'analysis' ? 'active' : '']" @click="activeTab = 'analysis'">{{
              $t('analysis') }}</button>
            <button :class="['tab-btn', activeTab === 'templates' ? 'active' : '']" @click="activeTab = 'templates'">{{
              $t('templates') }}</button>
          </div>
          <div style="margin-left:auto; font-size:13px" class="muted">{{ $t('currentFile') }} {{ fileName ||
            $t('noFile') }}</div>
        </div>

        <div style="padding:16px">
          <div v-show="activeTab === 'overview'">
            <h4 style="margin-top:0">{{ $t('dataPreview') }}</h4>
            <DataTable :data="tableData" @updateCell="onUpdateCell" @updateData="onUpdateData" :showExport="true" @export="onExport" />
          </div>

          <div v-show="activeTab === 'analysis'">
            <Analysis :saved-file-id="lastSavedFileId" :last-file="lastFile" />
          </div>

          <div v-show="activeTab === 'templates'">
            <Templates @template-response="onTemplateResponse" />
          </div>
        </div>
      </section>
      <aside class="right" style="width:420px; display:flex; flex-direction:column; gap:12px; height:100%;">
        <div class="card" style="display:flex; flex-direction:column; gap:12px;">
          <!-- Chat box: uploader (hidden after run), AI stream + steps (shown after run), and command input -->
          <div class="chat-box" style="display:flex; flex-direction:column; gap:8px;">
            <!-- Uploader: hidden when aiActive -->
            <div v-if="!aiActive">
              <FileUploader @fileLoaded="onFileLoaded" />
            </div>

            <!-- AI Stream and Steps: shown when aiActive -->
            <div v-if="aiActive" style="display:flex; flex-direction:column; gap:8px;">
                <div class="ai-stream-container">
                  <ChatBubbleList :messages="aiMessages" :executingIds="executingIds" :appliedIds="appliedIds" @apply-token="handleApplyToken" @skip-token="handleSkipToken" />
                </div>
              <!-- steps are included inside aiMessages as a single card message -->
            </div>

            <!-- Command input removed from chat box; moved to floating input at bottom-right -->
          </div>
        </div>
        <!-- Floating AI command input (aligned inside right sidebar) -->
        <div class="floating-ai-input">
          <ChatInput ref="aiInput" @send="onRunCommand" />
        </div>
      </aside>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted, watch } from 'vue'
import * as XLSX from 'xlsx'
import { useI18n } from 'vue-i18n'
import FileUploader from '../components/FileUploader.vue'
import DataTable from '../components/DataTable.vue'
import ChatBubbleList from '../components/ChatBubbleList.vue'
import ChatInput from '../components/ChatInput.vue'
import Analysis from './Analysis.vue'
import Templates from './Templates.vue'
import { getExcelDataPreview, processExcelWithAI, processExcelAndChat, uploadExcel, processExcelWithAIDownload, chat, runAllApis } from '../services/aiService'
import { handleUserChat } from '../services/chatManager'

const fileName = ref('')
const tableData = ref<Array<string[]>>([])
const aiMessages = ref<Array<{ id: number; role: string; text: string; placeholder?: boolean }>>([])
const aiActive = ref(false)
const executingIds = ref<string[]>([])
const appliedIds = ref<string[]>([])
const lastFile = ref<File | null>(null)
const lastCommand = ref<string>('')
const lastError = ref<string | null>(null)
const failedCommand = ref<string | null>(null)
// history of past chat interactions
const chatHistory = ref<Array<{ id: number, command: string, timestamp: number, messages: string[], ignored?: string }>>([])
const { t } = useI18n()
const aiInput = ref<any>(null)
const activeTab = ref('overview')
const autoSaveEnabled = ref(true)
const AUTO_SAVE_DEBOUNCE_MS = 2000
let autosaveTimer: number | null = null
const lastSavedFileId = ref<string | null>(null)

function onTemplateResponse(text: string) {
  aiActive.value = true
  const msg = String(text || '') || t('serverError')
  aiMessages.value.push({ id: Date.now(), role: 'ai', text: msg })
}

async function uploadTableToServer(wbout: Uint8Array | ArrayBuffer, name: string) {
  try {
    const blob = new Blob([wbout as any], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const form = new FormData()
    form.append('file', new File([blob], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
    const resp = await fetch('/api/excel/save', { method: 'POST', body: form })
    if (!resp.ok) throw new Error('upload failed')
    const j = await resp.json().catch(() => null)
    if (j && j.fileId) {
      lastSavedFileId.value = j.fileId
    }
    return j
  } catch (e) {
    return Promise.reject(e)
  }
}

function scheduleAutoSave() {
  try {
    if (autosaveTimer) clearTimeout(autosaveTimer)
    autosaveTimer = window.setTimeout(() => {
      try {
        const aoa = tableData.value || []
        const ws = XLSX.utils.aoa_to_sheet(aoa)
        const wb = XLSX.utils.book_new()
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
        const name = fileName.value ? `autosaved_${fileName.value}` : `autosave_${Date.now()}.xlsx`
        uploadTableToServer(wbout, name).catch(e => console.warn('autosave upload failed', e))
      } catch (e) { console.warn('autosave failed', e) }
    }, AUTO_SAVE_DEBOUNCE_MS) as unknown as number
  } catch (e) { console.warn('scheduleAutoSave error', e) }
}

async function onFileLoaded(payload: { name: string; data: string[][]; file?: File }) {
  fileName.value = payload.name
  tableData.value = payload.data
  aiMessages.value = []
  aiActive.value = false
  // store raw file if available for backend operations
  if (payload.file) {
    lastFile.value = payload.file
    // try server-side preview; fall back to client-parsed data on error
    try {
      const resp = await getExcelDataPreview(payload.file)
      // server returns parsed data under `data` or `excelDataPreview`; accept both
      if (resp && resp.data) tableData.value = resp.data as string[][]
      else if (resp && resp.excelDataPreview) tableData.value = resp.excelDataPreview as string[][]
      else tableData.value = payload.data
      } catch (e) {
      // ignore and keep client preview
      tableData.value = payload.data
    }
    // NOTE: 不在文件加载时自动触发所有 AI 接口（避免误调用如 suggest-charts）。
    // 如果需要针对性测试，请在 Settings 中启用或手动调用 runAllApis。
  } else {
    lastFile.value = null
  }
  saveToStorage()
}

async function onRunCommand(command: string) {
  // activate AI chat view
  aiActive.value = true
  lastError.value = null
  failedCommand.value = null
  lastCommand.value = command
  // snapshot of current messages for history
  chatHistory.value.push({ id: Date.now(), command, timestamp: Date.now(), messages: aiMessages.value.map(m => m.text) })

  // helper callbacks passed into chatManager
  function pushUserMessage(msg: { id: number; role: string; text: string; placeholder?: boolean }) {
    aiMessages.value.push(msg)
  }

  function pushAiPlaceholder(msg: { id: number; role: string; text: string; placeholder?: boolean }) {
    aiMessages.value.push(msg)
    return msg.id
  }

  function replaceAiMessage(placeholderId: number, newMsg: { id?: number; role?: string; text?: string }) {
    const idx = aiMessages.value.findIndex(m => m.id === placeholderId)
    if (idx !== -1) {
      aiMessages.value.splice(idx, 1, { id: newMsg.id || Date.now(), role: newMsg.role || 'ai', text: newMsg.text || '' })
    } else {
      aiMessages.value.push({ id: newMsg.id || Date.now(), role: newMsg.role || 'ai', text: newMsg.text || '' })
    }
  }

  function setTablePreview(data: string[][]) {
    try { tableData.value = data } catch (e) { /* ignore */ }
  }

  try {
    await handleUserChat({
      command,
      lastFile: lastFile.value,
      lastSavedFileId: lastSavedFileId.value,
      pushUserMessage: pushUserMessage as any,
      pushAiPlaceholder: pushAiPlaceholder as any,
      replaceAiMessage: replaceAiMessage as any,
      setTablePreview
    })
  } catch (err:any) {
    lastError.value = (err && (err.body?.error || err.message)) || t('serverError')
    failedCommand.value = command
  }
}

async function onExport() {
  if (!lastFile.value || !lastCommand.value) {
    alert(t('exportGuard'))
    return
  }
  try {
    const blob = await processExcelWithAIDownload(lastFile.value, lastCommand.value, lastSavedFileId.value || undefined)
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    const name = fileName.value ? `modified_${fileName.value}` : 'modified.xlsx'
    a.download = name
    document.body.appendChild(a)
    a.click()
    a.remove()
    URL.revokeObjectURL(url)
  } catch (e:any) {
    const msg = (e && (e.body?.error || e.message)) || t('exportFailed')
    lastError.value = msg
    failedCommand.value = lastCommand.value
    alert(msg)
  }
}

function retryLastCommand(){
  if (failedCommand.value) onRunCommand(failedCommand.value)
}

function newChat() {
  // start a fresh chat: clear messages and deactivate AI view
  aiActive.value = false
  aiMessages.value = []
  // clear loaded excel file and table data, persist change
  fileName.value = ''
  tableData.value = []
  saveToStorage()
  // record this new chat in history as an empty session marker
  chatHistory.value.push({ id: Date.now(), command: '', timestamp: Date.now(), messages: [] })
}

function onEditCommand(m: string) {
  // open AI view and populate the floating input for editing
  aiActive.value = true
  // set the input text to the command/message (strip UI labels if needed)
  const content = String(m || '')
  // try to set via exposed method
  if (aiInput.value && typeof aiInput.value.setText === 'function') {
    aiInput.value.setText(content)
  }
}

function onIgnoreCommand(m: any) {
  // remove message by text or id
  let idx = -1
  if (typeof m === 'number') idx = aiMessages.value.findIndex(x => x.id === m)
  else if (typeof m === 'string') idx = aiMessages.value.findIndex(x => x.text === m)
  else if (m && typeof m.id === 'number') idx = aiMessages.value.findIndex(x => x.id === m.id)
  if (idx !== -1) aiMessages.value.splice(idx, 1)
  // record ignored message in history
  chatHistory.value.push({ id: Date.now(), command: '', timestamp: Date.now(), messages: [], ignored: String(m && (m.id || m) || m) })
}

function onUpdateCell({ r, c, value }: { r: number, c: number, value: string }) {
  if (!tableData.value[r]) return
  tableData.value[r][c] = value
  saveToStorage()
}

function onUpdateData(newData: string[][]) { tableData.value = newData; saveToStorage() }

function saveToStorage() {
  try {
    const payload = { name: fileName.value, data: tableData.value }
    localStorage.setItem('aiexcel_workspace_file', JSON.stringify(payload))
  } catch (e) { }
}

// Watch table changes to trigger local storage save and debounced server autosave
watch(tableData, () => {
  saveToStorage()
  if (autoSaveEnabled.value) scheduleAutoSave()
}, { deep: true })

async function saveTableAsExcel() {
  // Create workbook and upload to backend storage without triggering a local download
  try {
    const aoa = tableData.value || []
    const ws = XLSX.utils.aoa_to_sheet(aoa)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
    const name = fileName.value ? `modified_${fileName.value}` : `modified_${Date.now()}.xlsx`
    try {
      const resp = await uploadTableToServer(wbout, name)
      return resp
    } catch (e) {
      console.warn('excel save/upload failed', e)
      throw e
    }
  } catch (e) {
    console.error('saveTableAsExcel error', e)
    throw e
  }
}

function loadFromStorage() {
  try {
    const s = localStorage.getItem('aiexcel_workspace_file')
    if (!s) return
    const p = JSON.parse(s)
    if (p && Array.isArray(p.data)) {
      fileName.value = p.name || ''
      tableData.value = p.data || []
    }
  } catch (e) { }
}

onMounted(() => loadFromStorage())

function colLetterToIndex(letters: string) { let n = 0; for (let i = 0; i < letters.length; i++) { n = n * 26 + (letters.charCodeAt(i) - 64) } return n - 1 }
function colIndexToLetter(idx: number) { let n = idx + 1; let s = ''; while (n > 0) { const rem = (n - 1) % 26; s = String.fromCharCode(65 + rem) + s; n = Math.floor((n - 1) / 26) } return s }
function evalFormula(expr: string, grid: string[][]) {
  const replaced = expr.replace(/([A-Z]+\d+)/g, (m) => {
    const col = colLetterToIndex(m.replace(/\d+/, ''))
    const row = Number(m.replace(/[^0-9]/g, ''))
    const v = (grid[row - 1] && grid[row - 1][col]) || '0'
    return String(Number(v) || 0)
  })
  try { // @ts-ignore
    return eval(replaced.replace(/[＝，]/g, ''))
  } catch (e) { return '' }
}

function getRangeValues(range: string) {
  const m = String(range || '').match(/^([A-Z]+\d+):([A-Z]+\d+)$/i)
  if (!m) return []
  const startRef = m[1]
  const endRef = m[2]
  const startCol = colLetterToIndex(startRef.replace(/\d+/, ''))
  const startRow = Number(startRef.replace(/[^0-9]/g, ''))
  const endCol = colLetterToIndex(endRef.replace(/\d+/, ''))
  const endRow = Number(endRef.replace(/[^0-9]/g, ''))
  const vals: Array<string | number> = []
  for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
    for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
      const row = tableData.value[r - 1]
      if (!row) continue
      vals.push(row[c])
    }
  }
  return vals
}

function parseNumber(v: any) {
  const num = Number(String(v || '').replace(/[^0-9.\-]/g, ''))
  return Number.isNaN(num) ? undefined : num
}

function countIf(values: Array<any>, condition: string) {
  const cond = String(condition || '').trim()
  const opMatch = cond.match(/^(>=|<=|>|<|==|=|!=)\s*(.+)$/)
  if (!opMatch) return 0
  const op = opMatch[1]
  const rhsRaw = opMatch[2]
  const rhsNum = Number(rhsRaw)
  return values.reduce((acc, v) => {
    const num = parseNumber(v)
    if (num === undefined) return acc
    switch (op) {
      case '>': return acc + (num > rhsNum ? 1 : 0)
      case '>=': return acc + (num >= rhsNum ? 1 : 0)
      case '<': return acc + (num < rhsNum ? 1 : 0)
      case '<=': return acc + (num <= rhsNum ? 1 : 0)
      case '!=': return acc + (num != rhsNum ? 1 : 0)
      case '==':
      case '=': return acc + (num === rhsNum ? 1 : 0)
      default: return acc
    }
  }, 0)
}

// Replace COUNTIF occurrences inside a larger expression so we can still eval arithmetic like COUNTIF(...)/36
function replaceCountIfExpressions(expr: string) {
  return String(expr || '').replace(/COUNTIF\(([A-Z]+\d+):([A-Z]+\d+),\s*"([^"]+)"\)/gi, (_, start: string, end: string, cond: string) => {
    const vals = getRangeValues(`${start}:${end}`)
    const cnt = countIf(vals, cond)
    return String(cnt)
  })
}

function ensureCell(ref: string) {
  const col = colLetterToIndex(ref.replace(/\d+/, ''))
  const row = Number(ref.replace(/[^0-9]/g, ''))
  while (tableData.value.length < row) {
    const cols = tableData.value[0] ? tableData.value[0].length : 0
    tableData.value.push(Array.from({ length: cols }, () => ''))
  }
  for (let r = 0; r < tableData.value.length; r++) {
    while ((tableData.value[r] || []).length <= col) {
      if (!tableData.value[r]) tableData.value[r] = []
      tableData.value[r].push('')
    }
  }
  return { rowIndex: row - 1, colIndex: col }
}

function ensureRows(count: number) {
  while (tableData.value.length < count) {
    const cols = tableData.value[0] ? tableData.value[0].length : 0
    tableData.value.push(Array.from({ length: cols }, () => ''))
  }
}

function adjustRowRefs(formula: string, fromRow: number, toRow: number) {
  const delta = toRow - fromRow
  return String(formula || '').replace(/([A-Z]+)(\d+)/g, (_, c: string, r: string) => {
    const nr = Number(r)
    if (Number.isNaN(nr)) return _
    return `${c}${nr + delta}`
  })
}

function applyCommands(raw: string) {
  const lines = raw.split('\n').map(l => l.trim()).filter(Boolean)
  for (const ln of lines) {
    let line = ln
    const colonMatch = line.match(/^(?:token:\s*)?([A-Z]+\d+)\s*:\s*(.+)$/i)
    if (colonMatch && !line.startsWith('[')) {
      // Normalize plain colon syntax to APPLY_FORMULA token
      line = `[APPLY_FORMULA:${colonMatch[1]}:${colonMatch[2]}]`
    }

    // FILL_DOWN by range e.g. J2:J35
    const fillPlainMatch = line.match(/^(?:token:\s*)?([A-Z]+\d+):([A-Z]+\d+)$/i)
    if (fillPlainMatch && !line.startsWith('[')) {
      line = `[FILL_DOWN:${fillPlainMatch[1]}:${fillPlainMatch[2]}]`
    }

    // Label + formula e.g. K36:总人数,=COUNTA(B2:B35) -> put label at K36, value at next column
    const labelFormulaMatch = line.match(/^(?:token:\s*)?([A-Z]+\d+):([^,]+),=?(.*)$/i)
    if (labelFormulaMatch && !line.startsWith('[')) {
      const ref = labelFormulaMatch[1]
      const label = labelFormulaMatch[2]
      const formula = labelFormulaMatch[3]
      const pos = ensureCell(ref)
      tableData.value[pos.rowIndex][pos.colIndex] = label
      const nextColRef = `${colIndexToLetter(pos.colIndex + 1)}${pos.rowIndex + 1}`
      applyCommands(`[APPLY_FORMULA:${nextColRef}:${formula}]`)
      continue
    }

    const mInsertCol = line.match(/\[INSERT_COLUMN:(\d+):([^\]]*)\]/i)
    if (mInsertCol) {
      const idx = Number(mInsertCol[1])
      const values = (mInsertCol[2] || '').split(',')
      const requiredRows = Math.max(tableData.value.length || 0, values.length || 0)
      ensureRows(requiredRows || 1)
      const colIndex = Math.max(0, Math.min(idx - 1, tableData.value[0] ? tableData.value[0].length : 0))
      for (let r = 0; r < tableData.value.length; r++) {
        const row = tableData.value[r]
        const val = values[r] !== undefined ? values[r] : ''
        row.splice(colIndex, 0, val)
      }
      continue
    }

    const mInsertRow = line.match(/\[INSERT_ROW:(\d+):([^\]]*)\]/i)
    if (mInsertRow) {
      const idx = Number(mInsertRow[1])
      const values = (mInsertRow[2] || '').split(',')
      const cols = Math.max(tableData.value[0] ? tableData.value[0].length : 0, values.length)
      if (!tableData.value[0]) tableData.value[0] = Array.from({ length: cols || 1 }, () => '')
      ensureRows(idx)
      const rowData = Array.from({ length: Math.max(cols, values.length, tableData.value[0].length) || 1 }, (_, i) => values[i] !== undefined ? values[i] : '')
      tableData.value.splice(Math.max(idx - 1, 0), 0, rowData)
      continue
    }

    const mDeleteRow = line.match(/\[DELETE_ROW:(\d+)\]/i)
    if (mDeleteRow) {
      const idx = Number(mDeleteRow[1])
      if (idx >= 1 && idx <= tableData.value.length) tableData.value.splice(idx - 1, 1)
      continue
    }

    const mDeleteCol = line.match(/\[DELETE_COLUMN:(\d+)\]/i)
    if (mDeleteCol) {
      const idx = Number(mDeleteCol[1]) - 1
      if (idx >= 0) {
        for (const row of tableData.value) {
          if (row && row.length > idx) row.splice(idx, 1)
        }
      }
      continue
    }

    const setMatch = line.match(/\[SET_CELL:([A-Z]+\d+):([^\]]+)\]/i)
    if (setMatch) {
      const ref = setMatch[1]
      const value = setMatch[2]
      const pos = ensureCell(ref)
      tableData.value[pos.rowIndex][pos.colIndex] = value
      continue
    }

    const fillMatch = line.match(/\[FILL_DOWN:([A-Z]+\d+):([A-Z]+\d+)\]/i)
    if (fillMatch) {
      const startRef = fillMatch[1]
      const endRef = fillMatch[2]
      const startPos = ensureCell(startRef)
      const endPos = ensureCell(endRef)
      const startRow = startPos.rowIndex
      const endRow = endPos.rowIndex
      const startColLetter = startRef.replace(/\d+/, '')
      const value = (tableData.value[startRow] && tableData.value[startRow][startPos.colIndex]) || ''
      for (let r = startRow; r <= endRow; r++) {
        const pos = ensureCell(`${startColLetter}${r + 1}`)
        const adjusted = adjustRowRefs(value, startRow + 1, r + 1)
        tableData.value[pos.rowIndex][startPos.colIndex] = adjusted
      }
      continue
    }
    // support both [APPLY_FORMULA:F2:=SUM(...)] and [APPLY_FORMULA:F2:SUM(...)]
    const m2 = line.match(/\[APPLY_FORMULA:([A-Z]+\d+):=?(.+)\]/i)
    if (m2) {
      const cellRef = m2[1]
      // strip optional leading '=' from expression
      let expr = m2[2].trim()
      if (expr.startsWith('=')) expr = expr.slice(1).trim()
      const targetCol = colLetterToIndex(cellRef.replace(/\d+/, ''))
      const targetRow = Number(cellRef.replace(/[^0-9]/g, ''))

      // Ensure table has enough rows
      while (tableData.value.length < targetRow) {
        // create empty row with same number of columns as header (or 0)
        const cols = tableData.value[0] ? tableData.value[0].length : 0
        const newRow = Array.from({ length: cols }, () => '')
        tableData.value.push(newRow)
      }
      // Ensure each row has enough columns
      for (let r = 0; r < tableData.value.length; r++) {
        while ((tableData.value[r] || []).length <= targetCol) {
          if (!tableData.value[r]) tableData.value[r] = []
          tableData.value[r].push('')
        }
      }

      // SUM(range)
      const sumMatch = expr.match(/^SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/i)
      if (sumMatch) {
        const vals = getRangeValues(`${sumMatch[1]}:${sumMatch[2]}`)
        const total = vals.reduce((acc: number, v) => {
          const num = parseNumber(v)
          return acc + (num !== undefined ? num : 0)
        }, 0)
        tableData.value[targetRow - 1][targetCol] = String(total)
        continue
      }

      // AVERAGE(range)
      const avgMatch = expr.match(/^AVERAGE\(([A-Z]+\d+):([A-Z]+\d+)\)$/i)
      if (avgMatch) {
        const vals = getRangeValues(`${avgMatch[1]}:${avgMatch[2]}`)
        const nums = vals.map(parseNumber).filter((v) => v !== undefined) as number[]
        const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0
        tableData.value[targetRow - 1][targetCol] = String(avg)
        continue
      }

      // IF(condition, trueVal, falseVal) simple compare against number
      const ifMatch = expr.match(/^IF\(([^,]+),([^,]+),(.+)\)$/i)
      if (ifMatch) {
        const condition = ifMatch[1].trim()
        const trueVal = ifMatch[2].trim()
        const falseVal = ifMatch[3].trim()
        const condMatch = condition.match(/([A-Z]+\d+)\s*(>=|<=|>|<|==|=|!=)\s*([0-9.]+)/i)
        let condResult = false
        if (condMatch) {
          const ref = condMatch[1]
          const op = condMatch[2]
          const rhs = Number(condMatch[3])
          const { rowIndex, colIndex } = ensureCell(ref)
          const cellVal = parseNumber(tableData.value[rowIndex][colIndex]) || 0
          switch (op) {
            case '>': condResult = cellVal > rhs; break
            case '>=': condResult = cellVal >= rhs; break
            case '<': condResult = cellVal < rhs; break
            case '<=': condResult = cellVal <= rhs; break
            case '!=': condResult = cellVal != rhs; break
            case '==':
            case '=': condResult = cellVal === rhs; break
          }
        }
        const pick = condResult ? trueVal : falseVal
        const cleaned = pick.replace(/^"|"$/g, '')
        tableData.value[targetRow - 1][targetCol] = cleaned
        continue
      }

      // STDEV.P / STDEV.S
      const stdevMatch = expr.match(/^STDEV\.(P|S)\(([A-Z]+\d+):([A-Z]+\d+)\)$/i)
      if (stdevMatch) {
        const vals = getRangeValues(`${stdevMatch[2]}:${stdevMatch[3]}`)
        const nums = vals.map(parseNumber).filter((v) => v !== undefined) as number[]
        const mean = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0
        const variance = nums.length ? nums.reduce((acc, v) => acc + Math.pow(v - mean, 2), 0) / (stdevMatch[1].toUpperCase() === 'S' && nums.length > 1 ? (nums.length - 1) : nums.length || 1) : 0
        const stdev = Math.sqrt(variance)
        tableData.value[targetRow - 1][targetCol] = String(stdev)
        continue
      }

      // COUNT(range) counts numeric values
      const countMatch = expr.match(/^COUNT\(([A-Z]+\d+):([A-Z]+\d+)\)$/i)
      if (countMatch) {
        const vals = getRangeValues(`${countMatch[1]}:${countMatch[2]}`)
        const cnt = vals.reduce((acc: number, v) => acc + (parseNumber(v) !== undefined ? 1 : 0), 0)
        tableData.value[targetRow - 1][targetCol] = String(cnt)
        continue
      }

      // COUNTA(range) counts non-empty
      const countAMatch = expr.match(/^COUNTA\(([A-Z]+\d+):([A-Z]+\d+)\)$/i)
      if (countAMatch) {
        const vals = getRangeValues(`${countAMatch[1]}:${countAMatch[2]}`)
        const cnt = vals.reduce((acc: number, v) => acc + (v !== undefined && v !== null && String(v).trim() !== '' ? 1 : 0), 0)
        tableData.value[targetRow - 1][targetCol] = String(cnt)
        continue
      }

      // COUNTIF(range, "<60")
      const countIfMatch = expr.match(/^COUNTIF\(([A-Z]+\d+):([A-Z]+\d+),\s*"([^"]+)"\)$/i)
      if (countIfMatch) {
        const vals = getRangeValues(`${countIfMatch[1]}:${countIfMatch[2]}`)
        const cond = countIfMatch[3]
        const cnt = countIf(vals, cond)
        tableData.value[targetRow - 1][targetCol] = String(cnt)
        continue
      }

      // Fallback: evaluate simple expressions by replacing cell refs with numeric values
      const exprWithCountIf = replaceCountIfExpressions(expr)
      const value = evalFormula(exprWithCountIf, tableData.value)
      if (tableData.value[targetRow - 1]) tableData.value[targetRow - 1][targetCol] = String(value)
      continue
    }
  }
  saveToStorage()
}

function onApplyCommands(m: string) { applyCommands(m) }

async function handleApply(cmd: string, id?: number) {
  const msgId = id || Date.now()
  const msgKey = String(msgId)
  if (!executingIds.value.includes(msgKey)) executingIds.value.push(msgKey)
  if (!appliedIds.value.includes(msgKey)) appliedIds.value.push(msgKey)
  try {
    const full = String(cmd || '')
    // find all APPLY_FORMULA tokens in order
    const tokenRegex = /(\[(?:APPLY_FORMULA|FILL_DOWN|SET_CELL|INSERT_ROW|INSERT_COLUMN|DELETE_ROW|DELETE_COLUMN):[^\]]+\])/ig
    const tokens = Array.from(full.matchAll(tokenRegex)).map(m => m[1])

    // helper: find nearest sentence containing token, fallback to full message
    function findSentenceForToken(message: string, token: string) {
      const parts = message.split(/(?<=。|\.|\n)/g).map(s => s.trim()).filter(Boolean)
      for (const p of parts) {
        if (p.indexOf(token) !== -1) return p
      }
      return message
    }

    const origIdx = aiMessages.value.findIndex(x => x.id === msgId)
    if (tokens.length > 1) {
      // Output word/token pairs in order, then execute each — merge into original AI message when possible
      for (const tok of tokens) {
        const inner = (tok.match(/\[(?:APPLY_FORMULA|FILL_DOWN|SET_CELL|INSERT_ROW|INSERT_COLUMN|DELETE_ROW|DELETE_COLUMN):([^\]]+)\]/i) || [])[1] || tok
        let word = findSentenceForToken(full, tok)
        // remove the bracketed token from displayed word to avoid re-parsing
        word = word.replace(tok, '').trim()
        const wordLine = `word: ${word}`
        const tokenLine = `token: ${inner}`

        if (origIdx !== -1) {
          const prev = aiMessages.value[origIdx].text || ''
          aiMessages.value.splice(origIdx, 1, { ...aiMessages.value[origIdx], text: `${prev}\n\n${wordLine}\n${tokenLine}` })
        } else {
          aiMessages.value.push({ id: Date.now(), role: 'ai', text: `${wordLine}\n${tokenLine}` })
        }

        await new Promise(res => setTimeout(res, 200))

        // execute token (applyCommands expects bracketed form, so pass tok)
        try {
          applyCommands(tok)
          try { await saveTableAsExcel() } catch (e) { /* ignore save errors */ }
          const doneLine = t('executedWith', { content: inner })
          if (origIdx !== -1) {
            const prev2 = aiMessages.value[origIdx].text || ''
            aiMessages.value.splice(origIdx, 1, { ...aiMessages.value[origIdx], text: `${prev2}\n${doneLine}` })
          } else {
            aiMessages.value.push({ id: Date.now(), role: 'ai', text: doneLine })
          }
        } catch (e:any) {
          const failLine = t('executeFailed', { content: inner, error: (e && e.message) || e })
          if (origIdx !== -1) {
            const prev2 = aiMessages.value[origIdx].text || ''
            aiMessages.value.splice(origIdx, 1, { ...aiMessages.value[origIdx], text: `${prev2}\n${failLine}` })
          } else {
            aiMessages.value.push({ id: Date.now(), role: 'ai', text: failLine })
          }
        }

        // small gap between tokens
        await new Promise(res => setTimeout(res, 300))
      }
    } else {
      // single or no token: fallback to prior behavior but merge outputs into original AI message
      const lines = full.split('\n').map(l => l.trim()).filter(Boolean)
      for (const ln of lines) {
        await new Promise(res => setTimeout(res, 300))
        try {
          applyCommands(ln)
          const doneLine = t('executedWith', { content: ln.replace(/\[|\]/g, '') })
          if (origIdx !== -1) {
            const prev = aiMessages.value[origIdx].text || ''
            aiMessages.value.splice(origIdx, 1, { ...aiMessages.value[origIdx], text: `${prev}\n${doneLine}` })
          } else {
            aiMessages.value.push({ id: Date.now(), role: 'ai', text: doneLine })
          }
        } catch (e:any) {
          const failLine = t('executeFailed', { content: ln, error: (e && e.message) || e })
          if (origIdx !== -1) {
            const prev = aiMessages.value[origIdx].text || ''
            aiMessages.value.splice(origIdx, 1, { ...aiMessages.value[origIdx], text: `${prev}\n${failLine}` })
          } else {
            aiMessages.value.push({ id: Date.now(), role: 'ai', text: failLine })
          }
        }
      }
    }
  } finally {
    const idx = executingIds.value.indexOf(msgKey)
    if (idx !== -1) executingIds.value.splice(idx, 1)
    // After executing, mark token-level as skipped/applied so UI no longer shows Execute button
    try { handleSkipToken(msgKey, msgId, undefined, t('executedNote')) } catch (e) { /* ignore */ }
  }
}

async function handleApplyToken(token: string, tokenKey: string, msgId?: number, idx?: number) {
  if (!tokenKey) return
  if (!executingIds.value.includes(tokenKey)) executingIds.value.push(tokenKey)
  if (!appliedIds.value.includes(tokenKey)) appliedIds.value.push(tokenKey)
  const origMsgId = typeof msgId === 'number' ? msgId : Number((tokenKey || '').split('-')[0])
  const origIdx = aiMessages.value.findIndex(x => x.id === origMsgId)
  const inner = (token.match(/\[(?:APPLY_FORMULA|FILL_DOWN|SET_CELL|INSERT_ROW|INSERT_COLUMN|DELETE_ROW|DELETE_COLUMN):([^\]]+)\]/i) || [])[1] || token

  // helper to find sentence
  function findSentenceForToken(message: string, tokenStr: string) {
    const parts = message.split(/(?<=。|\.|\n)/g).map(s => s.trim()).filter(Boolean)
    for (const p of parts) if (p.indexOf(tokenStr) !== -1) return p
    return message
  }

  try {
    const full = aiMessages.value[origIdx] && aiMessages.value[origIdx].text ? String(aiMessages.value[origIdx].text) : ''
    let word = findSentenceForToken(full, token)
    word = word.replace(token, '').trim()
    const wordLine = `word: ${word}`
    const tokenLine = `token: ${inner}`

    if (origIdx !== -1) {
      const prev = aiMessages.value[origIdx].text || ''
      aiMessages.value.splice(origIdx, 1, { ...aiMessages.value[origIdx], text: `${prev}\n\n${wordLine}\n${tokenLine}` })
    } else {
      aiMessages.value.push({ id: Date.now(), role: 'ai', text: `${wordLine}\n${tokenLine}` })
    }

    await new Promise(res => setTimeout(res, 200))

    // execute the token (expects bracketed token)
    try {
      applyCommands(token)
      try { await saveTableAsExcel() } catch (e) { /* ignore save errors */ }
      const doneLine = `已执行: ${inner}`
      if (origIdx !== -1) {
        const prev2 = aiMessages.value[origIdx].text || ''
        aiMessages.value.splice(origIdx, 1, { ...aiMessages.value[origIdx], text: `${prev2}\n${doneLine}` })
      } else {
        aiMessages.value.push({ id: Date.now(), role: 'ai', text: doneLine })
      }
    } catch (e:any) {
      const failLine = `执行失败: ${inner} -> ${(e && e.message) || e}`
      if (origIdx !== -1) {
        const prev2 = aiMessages.value[origIdx].text || ''
        aiMessages.value.splice(origIdx, 1, { ...aiMessages.value[origIdx], text: `${prev2}\n${failLine}` })
      } else {
        aiMessages.value.push({ id: Date.now(), role: 'ai', text: failLine })
      }
    }
  } finally {
    const i = executingIds.value.indexOf(tokenKey)
    if (i !== -1) executingIds.value.splice(i, 1)
  }
}

function handleSkipToken(tokenKey: string, msgId?: number, idx?: number, note?: string) {
  if (!tokenKey) return
  if (!appliedIds.value.includes(tokenKey)) appliedIds.value.push(tokenKey)
  const origMsgId = typeof msgId === 'number' ? msgId : Number((tokenKey || '').split('-')[0])
  const origIdx = aiMessages.value.findIndex(x => x.id === origMsgId)
  const text = note || t('skippedNote')
  if (origIdx !== -1) {
    const prev = aiMessages.value[origIdx].text || ''
    aiMessages.value.splice(origIdx, 1, { ...aiMessages.value[origIdx], text: `${prev}\n${text}` })
  } else {
    aiMessages.value.push({ id: Date.now(), role: 'ai', text })
  }
}
</script>

<style scoped>
/* grid to ensure topbar and preview share the same main column width */
.workspace-grid {
  display: grid;
  grid-template-columns: 1fr 420px;
  grid-template-rows: auto 1fr;
  gap: 12px;
  align-items: start;
  height: calc(100vh - 32px);
  max-height: calc(100vh - 32px);
  min-width: 0;
  min-height: 0;
  overflow: hidden;
}

.workspace-grid>.topbar {
  grid-column: 1 / 2;
  grid-row: 1;
}

.workspace-grid>.preview {
  grid-column: 1 / 2;
  grid-row: 2;
  height: 100%;
  min-height: 0;
}

.workspace-grid>.right {
  grid-column: 2 / 3;
  grid-row: 1 / 3;
  height: 100%;
  min-height: 0;
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.workspace-grid>.right .card {
  flex: 1;
  min-height: 0;
  overflow: hidden;
}

.workspace-grid>.right .chat-box {
  flex: 1;
  min-height: 0;
  overflow: auto;
}

.tab-btn {
  background: transparent;
  border: 1px solid transparent;
  padding: 6px 12px;
  border-radius: 6px;
  cursor: pointer;
  color: #cfcfd6;
}

.tab-btn.active {
  background: #189079;
  /* workspace accent color */
  border-color: transparent;
  color: #fff;
}

.workspace-tabs {
  border-bottom: 1px solid #eee
}

.tab-btn:hover {
  filter: brightness(1.05);
}
</style>
