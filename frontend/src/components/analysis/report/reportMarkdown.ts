function escapeHtml(text: string) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function formatInline(text: string) {
  return escapeHtml(text)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.+?)\*/g, '<em>$1</em>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank" rel="noopener noreferrer">$1</a>')
}

export function renderMarkdown(source: string) {
  const lines = String(source || '').replace(/\r\n/g, '\n').split('\n')
  const html: string[] = []
  let listType: 'ul' | 'ol' | null = null
  let listItems: string[] = []

  const flushList = () => {
    if (!listType || listItems.length === 0) return
    html.push(`<${listType}>${listItems.map(item => `<li>${item}</li>`).join('')}</${listType}>`)
    listType = null
    listItems = []
  }

  for (const rawLine of lines) {
    const line = rawLine.trim()
    if (!line) {
      flushList()
      continue
    }

    const heading = line.match(/^(#{1,6})\s+(.*)$/)
    if (heading) {
      flushList()
      const level = heading[1].length
      html.push(`<h${level}>${formatInline(heading[2])}</h${level}>`)
      continue
    }

    const unordered = line.match(/^[-*+]\s+(.*)$/)
    if (unordered) {
      if (listType !== 'ul') {
        flushList()
        listType = 'ul'
      }
      listItems.push(formatInline(unordered[1]))
      continue
    }

    const ordered = line.match(/^\d+\.\s+(.*)$/)
    if (ordered) {
      if (listType !== 'ol') {
        flushList()
        listType = 'ol'
      }
      listItems.push(formatInline(ordered[1]))
      continue
    }

    flushList()
    html.push(`<p>${formatInline(line)}</p>`)
  }

  flushList()
  return html.join('\n')
}