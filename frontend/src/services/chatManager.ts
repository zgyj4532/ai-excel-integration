import { processExcelAndChat, chat, analyzeExcelData, createChart } from './aiService'

type PushMsgFn = (m: { id: number; role: string; text: string; placeholder?: boolean }) => void
type ReplaceMsgFn = (placeholderId: number, msg: { id?: number; role?: string; text?: string }) => void
type SetTableFn = (data: string[][]) => void

export const apiActions = {
  excelWithAi: async (file: File | null, command: string, fileId?: string | undefined) => {
    return await processExcelAndChat(file, command, fileId)
  },
  chat: async (message: string) => {
    return await chat(message)
  }
}

// extend apiActions with analysis and chart creation
apiActions['analyzeData'] = async (file: File | null, analysisRequest: string, fileId?: string | undefined) => {
  // prefer sending file if provided and fileId absent
  return await analyzeExcelData(file as any, analysisRequest)
}
apiActions['createChart'] = async (file: File | null, chartType: string, targetColumn: string, fileId?: string | undefined) => {
  return await createChart(file as any, chartType, targetColumn, fileId)
}

/**
 * handleUserChat: a small adapter that wraps aiService calls and reports progress
 * opts:
 *  - command: user command
 *  - lastFile: optional File
 *  - lastSavedFileId: optional fileId string
 *  - pushUserMessage: pushes a user message into UI
 *  - pushAiPlaceholder: pushes an AI placeholder and returns its id
 *  - replaceAiMessage: replaces placeholder with final AI message
 *  - setTablePreview: optional table data updater
 */
export async function handleUserChat(opts: {
  command: string
  lastFile?: File | null
  lastSavedFileId?: string | null
  pushUserMessage: PushMsgFn
  pushAiPlaceholder: (msg: { id: number; role: string; text: string; placeholder?: boolean }) => number
  replaceAiMessage: ReplaceMsgFn
  setTablePreview?: SetTableFn
}) {
  const { command, lastFile, lastSavedFileId, pushUserMessage, pushAiPlaceholder, replaceAiMessage, setTablePreview } = opts

  // push user message
  pushUserMessage({ id: Date.now(), role: 'user', text: command })

  // decide whether to call excel+ai endpoint or plain chat
  if (lastFile || lastSavedFileId) {
    const placeholder = { id: Date.now() + Math.floor(Math.random() * 1000), role: 'ai', text: '', placeholder: true }
    const placeholderId = pushAiPlaceholder(placeholder)
    try {
      const { preview, aiResp } = await apiActions.excelWithAi(lastFile || null, command, lastSavedFileId || undefined)
      if (preview && setTablePreview) {
        if (Array.isArray(preview.data)) setTablePreview(preview.data as string[][])
        else if (preview.excelDataPreview && typeof preview.excelDataPreview === 'string') {
          try {
            const maybe = JSON.parse(preview.excelDataPreview)
            if (Array.isArray(maybe)) setTablePreview(maybe as string[][])
          } catch (e) {}
        }
      }

      const aiText = (aiResp && (aiResp.aiResponse || aiResp.message || aiResp.result || JSON.stringify(aiResp))) as string
      replaceAiMessage(placeholderId, { id: Date.now(), role: 'ai', text: aiText })
      return { preview, aiResp }
    } catch (err:any) {
      const msg = (err && (err.body?.error || err.message)) || '服务端错误'
      // propagate the failure as replacement message
      replaceAiMessage(placeholderId, { id: Date.now(), role: 'ai', text: msg })
      throw err
    }
  } else {
    const placeholder = { id: Date.now() + Math.floor(Math.random() * 1000), role: 'ai', text: '', placeholder: true }
    const placeholderId = pushAiPlaceholder(placeholder)
    try {
      const resp = await apiActions.chat(command)
      const text = (resp && (resp.message || resp.result || JSON.stringify(resp))) as string
      replaceAiMessage(placeholderId, { id: Date.now(), role: 'ai', text })
      return { aiResp: resp }
    } catch (err:any) {
      const msg = (err && (err.body?.error || err.message)) || '服务端错误'
      replaceAiMessage(placeholderId, { id: Date.now(), role: 'ai', text: msg })
      throw err
    }
  }
}

// Additional helper: process inline tool tokens inside an AI response text.
// This will find known tokens and execute them via apiActions, appending progress messages via provided callbacks.
export async function processAiToolTokens(text: string, opts: {
  lastFile?: File | null
  lastSavedFileId?: string | null
  pushProgress: (line: string) => void
}) {
  // ANALYZE_DATA: [ANALYZE_DATA:request]
  const analyzeRegex = /\[ANALYZE_DATA:([^\]]+)\]/ig
  const createChartRegex = /\[CREATE_CHART:([^:\]]+):([^\]]+)\]/ig

  // process ANALYZE_DATA tokens sequentially
  for (const m of Array.from(text.matchAll(analyzeRegex))) {
    const req = (m[1] || '').trim()
    pushProgress(`开始分析: ${req}`)
    try {
      const res = await apiActions['analyzeData'](opts.lastFile || null, req, opts.lastSavedFileId || undefined)
      pushProgress(`分析完成: ${req}`)
    } catch (e:any) {
      pushProgress(`分析失败: ${req} -> ${(e && e.message) || e}`)
    }
    await new Promise(r => setTimeout(r, 200))
  }

  // process CREATE_CHART tokens sequentially
  for (const m of Array.from(text.matchAll(createChartRegex))) {
    const chartType = (m[1] || '').trim()
    const col = (m[2] || '').trim()
    pushProgress(`创建图表: ${chartType} 列: ${col}`)
    try {
      const res = await apiActions['createChart'](opts.lastFile || null, chartType, col, opts.lastSavedFileId || undefined)
      pushProgress(`图表创建成功: ${chartType} ${col}`)
    } catch (e:any) {
      pushProgress(`图表创建失败: ${chartType} ${col} -> ${(e && e.message) || e}`)
    }
    await new Promise(r => setTimeout(r, 200))
  }
}

export default { apiActions, handleUserChat }
