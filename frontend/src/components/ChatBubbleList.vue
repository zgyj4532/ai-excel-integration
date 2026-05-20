<template>
  <div class="chat-bubble-list">
    <div v-for="m in messages" :key="m.id" :class="['bubble-row', m.role === 'user' ? 'user' : 'ai']">
      <div class="bubble" :class="{placeholder: m.placeholder}">
        <div v-if="m.placeholder" class="placeholder-dot">
          <span></span><span></span><span></span>
        </div>
        <div v-else>
          <template v-if="hasApply(m.text)">
            <div>
              <div v-for="(seg, i) in splitAllApply(m.text)" :key="i" style="width:100%;">
                <div v-if="seg.type === 'text'" v-html="escapeHtml(seg.text)"></div>
                <div v-else-if="seg.type === 'token'">
                  <div class="apply-text applied">{{ seg.inner }}</div>
                </div>
              </div>
            </div>
          </template>
          <template v-else>
            <div v-html="escapeHtml(m.text)"></div>
          </template>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { watch } from 'vue'

type SnapshotOps = {
  getActiveWorkbook?: () => any
  getFormulaEngine?: () => any
}

const props = defineProps<{ messages: Array<{ id: number; role: string; text: string; placeholder?: boolean }>, executingIds?: string[], appliedIds?: string[], snapshotOps?: SnapshotOps | null }>()
const emit = defineEmits(['apply-token','skip-token'])

const processedMessages = new Set<string>()
const ENABLE_AUTO_EXCEL_APPLY = true

watch(
  () => props.messages.map(m => ({ id: m.id, role: m.role, text: m.text, placeholder: m.placeholder })),
  (messages) => {
    if (!ENABLE_AUTO_EXCEL_APPLY) return
    messages.forEach(({ id, role, text, placeholder }) => {
      if (role !== 'ai' || placeholder) return
      const key = `${id}:${text}`
      if (processedMessages.has(key)) return
      const success = applyExcelCommandsFromMessage(text)
      if (success) {
        processedMessages.add(key)
      }
    })
  },
  { deep: true, immediate: true }
)

watch(
  () => props.snapshotOps,
  () => {
    if (!ENABLE_AUTO_EXCEL_APPLY) return
    props.messages.forEach((m) => {
      if (m.role !== 'ai' || m.placeholder) return
      const key = `${m.id}:${m.text}`
      if (processedMessages.has(key)) return
      const success = applyExcelCommandsFromMessage(m.text)
      if (success) {
        processedMessages.add(key)
      }
    })
  },
  { immediate: true }
)

function applyExcelCommandsFromMessage(raw: string) {
  if (!raw) return false
  const sheet = getActiveSheet()
  if (!sheet) return false

  const assignments = extractAssignments(raw)
  let applied = false
  assignments.forEach(({ columnLetters, rowIndex, value }) => {
    try {
      ensureColumnExists(sheet, columnLetters)
      ensureRowExists(sheet, rowIndex)
      const cellRef = `${columnLetters}${rowIndex + 1}`
      const normalizedValue = normalizeAssignmentValue(value)
      const range = sheet.getRange?.(cellRef) ?? sheet.getRangeByA1?.(cellRef)
      if (!range) return
      if (typeof range.setValue === 'function') {
        range.setValue(normalizedValue)
      } else if (typeof range.setValues === 'function') {
        range.setValues([[normalizedValue]])
      }
      applied = true
    } catch (error) {
      console.warn('Failed to apply assignment', error)
    }
  })
  if (applied) {
    try {
      const engine = props.snapshotOps?.getFormulaEngine?.()
      engine?.executeCalculation?.()
    } catch (error) {
      console.warn('Formula calculation failed', error)
    }
  }
  return applied
}

function getActiveSheet() {
  const workbook = props.snapshotOps?.getActiveWorkbook?.()
  if (!workbook || typeof workbook.getActiveSheet !== 'function') return null
  return workbook.getActiveSheet()
}

function extractAssignments(raw: string) {
  const results: Array<{ columnLetters: string; rowIndex: number; value: string }> = []
  const sanitized = String(raw || '')
    .replace(/token:\s*/gi, '')
    .replace(/：/g, ':')
    .replace(/，/g, ',')

  const assignmentHead = /([A-Z]+)(\d+)\s*(?:=|:)/gi
  let match: RegExpExecArray | null

  while ((match = assignmentHead.exec(sanitized))) {
    const columnLetters = match[1].toUpperCase()
    const rowNumber = Number(match[2])
    if (!rowNumber) continue
    const rowIndex = rowNumber - 1
    const valueStart = assignmentHead.lastIndex
    let valueEnd = sanitized.length
    for (let i = valueStart; i < sanitized.length; i++) {
      const char = sanitized[i]
      if (char === '\n' || char === ';') {
        valueEnd = i
        break
      }
    }
    const value = sanitized.slice(valueStart, valueEnd).trim()
    results.push({ columnLetters, rowIndex, value })
    assignmentHead.lastIndex = valueEnd
  }
  const applyFormula = /APPLY_FORMULA:([A-Z]+)(\d+):?=([^\]\n;]+)/gi
  while ((match = applyFormula.exec(sanitized))) {
    const columnLetters = match[1].toUpperCase()
    const rowNumber = Number(match[2])
    if (!rowNumber) continue
    const rowIndex = rowNumber - 1
    const value = match[3].trim()
    results.push({ columnLetters, rowIndex, value })
  }
  return results
}

function normalizeAssignmentValue(raw: string) {
  let text = String(raw || '').trim()
  text = text.replace(/[“”]/g, '"').replace(/[‘’]/g, "'")
  text = text.replace(/[。.;]+$/u, '').replace(/,+$/u, '').trim()
  // Some auto-generated tokens may trail with a stray ']' (e.g., APPLY_FORMULA blocks); strip it to avoid broken formulas
  text = text.replace(/\]+$/u, '').trim()
  if (text.startsWith('=')) {
    const body = text.slice(1).trim()
    const isQuoted = body.startsWith('"') || body.startsWith("'")
    const looksLikeFormula = /^[_A-Za-z][\w.]*\s*\(|^[A-Z]+\d+|^[-+]?\d/.test(body)
    if (body && !isQuoted && !looksLikeFormula) {
      const escaped = body.replace(/"/g, '""')
      return `="${escaped}"`
    }
  }
  if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
    const inner = text.slice(1, -1)
    const escaped = inner.replace(/"/g, '""')
    return `="${escaped}"`
  }
  if (!text.startsWith('=')) {
    return `= ${text}`
  }
  return text
}

function ensureColumnExists(sheet: any, columnLetters: string) {
  const targetIndex = columnLetterToIndex(columnLetters)
  if (targetIndex < 0) return
  let current = getColumnCount(sheet)
  if (current === 0) {
    if (typeof sheet.insertColumns === 'function') {
      sheet.insertColumns(0, targetIndex + 1)
      current = targetIndex + 1
    } else if (typeof sheet.insertColumnBefore === 'function') {
      for (let i = 0; i <= targetIndex; i++) {
        sheet.insertColumnBefore(0)
        current += 1
      }
    }
  }
  if (current === 0) current = 1
  while (current <= targetIndex) {
    const diff = targetIndex + 1 - current
    if (diff <= 0) break
    if (typeof sheet.insertColumns === 'function') {
      sheet.insertColumns(current, diff)
      current += diff
    } else if (typeof sheet.insertColumnAfter === 'function') {
      sheet.insertColumnAfter(current - 1)
      current += 1
    } else if (typeof sheet.insertColumnsAfter === 'function') {
      sheet.insertColumnsAfter(current - 1, diff)
      current += diff
    } else {
      break
    }
  }
}

function ensureRowExists(sheet: any, rowIndex: number) {
  if (rowIndex < 0) return
  let current = getRowCount(sheet)
  if (current === 0) {
    if (typeof sheet.insertRows === 'function') {
      sheet.insertRows(0, rowIndex + 1)
      current = rowIndex + 1
    } else if (typeof sheet.insertRowBefore === 'function') {
      for (let i = 0; i <= rowIndex; i++) {
        sheet.insertRowBefore(0)
        current += 1
      }
    }
  }
  if (current === 0) current = 1
  while (current <= rowIndex) {
    const diff = rowIndex + 1 - current
    if (diff <= 0) break
    if (typeof sheet.insertRows === 'function') {
      sheet.insertRows(current, diff)
      current += diff
    } else if (typeof sheet.insertRowAfter === 'function') {
      sheet.insertRowAfter(current - 1)
      current += 1
    } else if (typeof sheet.insertRowsAfter === 'function') {
      sheet.insertRowsAfter(current - 1, diff)
      current += diff
    } else {
      break
    }
  }
}

function getColumnCount(sheet: any) {
  if (!sheet) return 0
  if (typeof sheet.getColumnCount === 'function') return sheet.getColumnCount()
  if (typeof sheet.getColumnsCount === 'function') return sheet.getColumnsCount()
  if (typeof sheet.getMaxColumns === 'function') return sheet.getMaxColumns()
  return 0
}

function getRowCount(sheet: any) {
  if (!sheet) return 0
  if (typeof sheet.getRowCount === 'function') return sheet.getRowCount()
  if (typeof sheet.getRowsCount === 'function') return sheet.getRowsCount()
  if (typeof sheet.getMaxRows === 'function') return sheet.getMaxRows()
  return 0
}

function columnLetterToIndex(letters: string) {
  let result = 0
  const upper = letters.toUpperCase()
  for (let i = 0; i < upper.length; i++) {
    const charCode = upper.charCodeAt(i)
    if (charCode < 65 || charCode > 90) return -1
    result = result * 26 + (charCode - 64)
  }
  return result - 1
}

function tokenKey(msgId: number, idx: number) {
  return `${msgId}-${idx}`
}

function escapeHtml(s: string) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>')
}

function hasApply(s: string) {
  if (!s) return false
  return /(\[(APPLY_FORMULA|FILL_DOWN|SET_CELL|INSERT_ROW|INSERT_COLUMN|DELETE_ROW|DELETE_COLUMN):[^\]]+\])|(^|\n)(token:)?\s*[A-Z]+\d+\s*:/i.test(s)
}

function splitAllApply(s: string) {
  const lines = String(s || '').split(/\n/)
  const parts: Array<any> = []
  const bracketRe = /(\[(APPLY_FORMULA|FILL_DOWN|SET_CELL|INSERT_ROW|INSERT_COLUMN|DELETE_ROW|DELETE_COLUMN):[^\]]+\])/i
  const colonRe = /^(?:token:\s*)?([A-Z]+\d+)\s*:\s*(.+)$/i

  for (const ln of lines) {
    const trimmed = ln.trim()
    if (!trimmed) continue

    // bracketed tokens
    if (bracketRe.test(trimmed)) {
      const matches = trimmed.match(new RegExp(bracketRe, 'gi')) || []
      let cursor = trimmed
      for (const tok of matches) {
        const idx = cursor.indexOf(tok)
        const prefix = cursor.slice(0, idx)
        if (prefix && prefix.trim()) parts.push({ type: 'text', text: prefix })
        const innerMatch = tok.match(/\[(APPLY_FORMULA|FILL_DOWN|SET_CELL|INSERT_ROW|INSERT_COLUMN|DELETE_ROW|DELETE_COLUMN):([^\]]+)\]/i)
        const inner = innerMatch ? innerMatch[2] : tok
        parts.push({ type: 'token', token: tok, inner })
        cursor = cursor.slice(idx + tok.length)
      }
      if (cursor && cursor.trim()) parts.push({ type: 'text', text: cursor })
      continue
    }

    // colon pattern -> treat as APPLY_FORMULA token
    const cm = trimmed.match(colonRe)
    if (cm) {
      const ref = cm[1]
      const expr = cm[2]
      const tok = `[APPLY_FORMULA:${ref}:${expr}]`
      parts.push({ type: 'token', token: tok, inner: `${ref}: ${expr}` })
      continue
    }

    // fallback plain text
    parts.push({ type: 'text', text: trimmed })
  }
  return parts
}
</script>

<style scoped>
.chat-bubble-list { display:flex; flex-direction:column; gap:8px; padding:8px; max-height:100%; overflow:visible }
.bubble-row { display:flex }
.bubble-row.user { justify-content:flex-end }
.bubble-row.ai { justify-content:flex-start }
.bubble { max-width:78%; padding:10px 14px; border-radius:14px; font-size:14px; line-height:1.4 }
/* apply styles on the inner .bubble based on the parent row role */
.bubble-row.user .bubble {
  background: linear-gradient(90deg,#8a2be2,#6d04c4) !important;
  color: #fff !important;
  box-shadow: 0 6px 18px rgba(109,4,196,0.12) !important;
  border: 1px solid rgba(109,4,196,0.18) !important;
}
.bubble-row.ai .bubble {
  background: var(--panel) !important;
  color: var(--text-primary) !important;
}
.bubble.placeholder { background: rgba(255,255,255,0.04) !important; color:transparent !important }
.placeholder-dot { display:flex; gap:6px; align-items:center; }
.placeholder-dot span { width:8px; height:8px; background: var(--text-primary); border-radius:50%; opacity:0.18; animation: blink 1s infinite }
.placeholder-dot span:nth-child(2){ animation-delay:0.15s }
.placeholder-dot span:nth-child(3){ animation-delay:0.3s }
@keyframes blink{0%{opacity:0.18}50%{opacity:0.9}100%{opacity:0.18}}

.apply-box { display:flex; flex-direction:column; align-items:center; gap:8px; }
.apply-text { text-align:center; font-weight:600; padding:8px 12px; border-radius:8px; background: var(--panel); color: var(--text-primary); }
.apply-text.applied { opacity:0.9; background: var(--panel); }
.apply-actions { display:flex; gap:8px }
.btn-execute { background: linear-gradient(90deg,#178C76,#178C83); color:var(--text-on-accent); border:none; padding:6px 12px; border-radius:6px; cursor:pointer }
.btn-skip { background: #6b7280; color:var(--text-on-accent); border:none; padding:6px 12px; border-radius:6px; cursor:pointer }
</style>
