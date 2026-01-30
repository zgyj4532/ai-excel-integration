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
                  <div v-if="!(props.appliedIds || []).includes(tokenKey(m.id,i))" class="apply-box">
                    <div class="apply-text">{{ seg.inner }}</div>
                    <div class="apply-actions">
                      <button class="btn-execute" :disabled="(props.executingIds || []).includes(tokenKey(m.id,i))" @click="$emit('apply-token', seg.token, tokenKey(m.id,i), m.id, i)">
                        {{ (props.executingIds || []).includes(tokenKey(m.id,i)) ? '执行中...' : '执行' }}
                      </button>
                      <button class="btn-skip" @click="$emit('skip-token', tokenKey(m.id,i), m.id, i)">跳过</button>
                    </div>
                  </div>
                  <div v-else class="apply-text applied">{{ seg.inner }}</div>
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
const props = defineProps<{ messages: Array<{ id: number; role: string; text: string; placeholder?: boolean }>, executingIds?: string[], appliedIds?: string[] }>()
const emit = defineEmits(['apply-token','skip-token'])

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

function extractInner(s: string) {
  const m = String(s).match(/\[APPLY_FORMULA:([^\]]+)\]/i)
  return m ? m[1] : s
}

function extractFull(s: string) {
  const m = String(s).match(/(\[APPLY_FORMULA:[^\]]+\])/i)
  return m ? m[1] : s
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
.chat-bubble-list { display:flex; flex-direction:column; gap:8px; padding:8px; height:700px; overflow:auto }
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
  background: rgba(255,255,255,0.04) !important;
  color: #fff !important;
}
.bubble.placeholder { background: rgba(255,255,255,0.04) !important; color:transparent !important }
.placeholder-dot { display:flex; gap:6px; align-items:center; }
.placeholder-dot span { width:8px; height:8px; background: #fff; border-radius:50%; opacity:0.18; animation: blink 1s infinite }
.placeholder-dot span:nth-child(2){ animation-delay:0.15s }
.placeholder-dot span:nth-child(3){ animation-delay:0.3s }
@keyframes blink{0%{opacity:0.18}50%{opacity:0.9}100%{opacity:0.18}}

.apply-box { display:flex; flex-direction:column; align-items:center; gap:8px; }
.apply-text { text-align:center; font-weight:600; padding:8px 12px; border-radius:8px; background: rgba(255,255,255,0.02); }
.apply-text.applied { opacity:0.9; background: rgba(255,255,255,0.01); }
.apply-actions { display:flex; gap:8px }
.btn-execute { background: linear-gradient(90deg,#178C76,#178C83); color:#fff; border:none; padding:6px 12px; border-radius:6px; cursor:pointer }
.btn-skip { background: #6b7280; color:#fff; border:none; padding:6px 12px; border-radius:6px; cursor:pointer }
</style>
