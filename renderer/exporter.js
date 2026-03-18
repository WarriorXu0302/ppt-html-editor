/**
 * exporter.js — Export slides as PPTX, editable PPTX, PDF, PNG/JPEG
 *
 * Standard PPTX: html2canvas renders each slide → PNG → pptxgenjs full-slide image
 *   → pixel-perfect fidelity, not editable in PowerPoint
 *
 * Editable PPTX: DOM traversal extracts text elements + background color → pptxgenjs
 *   → native text boxes, editable in PowerPoint, ~80-90% visual fidelity
 */

let exportCancelled = false

// Reusable iframe to avoid DOM thrashing during batch exports
let cachedIframe = null

function getOrCreateIframe() {
  if (cachedIframe && cachedIframe.parentNode) return cachedIframe
  const iframe = document.createElement('iframe')
  iframe.style.cssText = [
    'position:fixed', 'left:-9999px', 'top:-9999px',
    'width:1280px', 'height:720px', 'border:none',
    'pointer-events:none', 'z-index:-1'
  ].join(';')
  iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin')
  document.body.appendChild(iframe)
  cachedIframe = iframe
  return iframe
}

function cleanupCachedIframe() {
  if (cachedIframe && cachedIframe.parentNode) {
    try { document.body.removeChild(cachedIframe) } catch (_) {}
  }
  cachedIframe = null
}

export function initExporter() {
  const modal = document.getElementById('export-modal')
  modal.querySelector('.modal-close').addEventListener('click', hideExportModal)
  document.getElementById('export-cancel-btn').addEventListener('click', () => {
    exportCancelled = true
    hideExportModal()
  })
  document.getElementById('export-confirm-btn').addEventListener('click', startExport)
}

export function showExportModal(totalSlides) {
  exportCancelled = false
  document.getElementById('export-total-hint').textContent = `(总共 ${totalSlides} 页)`
  document.getElementById('export-progress-section').style.display = 'none'
  document.getElementById('export-confirm-btn').disabled = false
  document.getElementById('export-confirm-btn').textContent = '开始导出'
  document.getElementById('export-modal').classList.remove('hidden')
}

function hideExportModal() {
  document.getElementById('export-modal').classList.add('hidden')
}

async function startExport() {
  const slides = window.appState?.slides
  if (!slides || slides.length === 0) return

  const format = document.querySelector('input[name="export-format"]:checked')?.value || 'pptx'
  const rangeType = document.querySelector('input[name="export-range"]:checked')?.value || 'all'
  const scale = parseFloat(document.querySelector('input[name="export-scale"]:checked')?.value || '1.5')
  const currentIndex = window.appState?.currentIndex || 0

  // Parse slide range
  let indices = []
  if (rangeType === 'all') {
    indices = slides.map((_, i) => i)
  } else if (rangeType === 'current') {
    indices = [currentIndex]
  } else {
    const rangeStr = document.getElementById('export-range-input').value
    indices = parseRange(rangeStr, slides.length)
    if (indices.length === 0) {
      alert('无效的范围，请输入如 "1-3, 5, 7" 格式')
      return
    }
  }

  // Show progress UI
  document.getElementById('export-progress-section').style.display = 'block'
  document.getElementById('export-confirm-btn').disabled = true
  document.getElementById('export-confirm-btn').textContent = '导出中...'
  exportCancelled = false

  try {
    if (format === 'pptx') {
      await exportPPTX(slides, indices, scale)
    } else if (format === 'editable-pptx') {
      await exportEditablePPTX(slides, indices)
    } else if (format === 'pdf') {
      await exportPDF(slides, indices, scale)
    } else {
      await exportImages(slides, indices, scale, format)
    }
    if (!exportCancelled) hideExportModal()
  } catch (err) {
    console.error('Export failed:', err)
    alert('导出失败：' + err.message)
  } finally {
    document.getElementById('export-confirm-btn').disabled = false
    document.getElementById('export-confirm-btn').textContent = '开始导出'
    document.getElementById('export-progress-section').style.display = 'none'
  }
}

// ── Core: render one slide to canvas ─────────────────────────────────────

async function renderSlideToCanvas(htmlContent, scale) {
  return new Promise((resolve, reject) => {
    const iframe = getOrCreateIframe()

    // Give the browser time to render after content loads
    const capture = () => {
      setTimeout(async () => {
        try {
          const canvas = await html2canvas(iframe.contentDocument.body, {
            scale,
            width: 1280,
            height: 720,
            useCORS: true,
            allowTaint: true,
            backgroundColor: '#ffffff',
            logging: false,
            windowWidth: 1280,
            windowHeight: 720
          })
          resolve(canvas)
        } catch (err) {
          reject(err)
        }
      }, 300)
    }

    // Load via blob URL so external resources (fonts, etc.) can still load
    const blob = new Blob([htmlContent], { type: 'text/html' })
    const blobUrl = URL.createObjectURL(blob)
    iframe.src = blobUrl
    iframe.onload = () => { URL.revokeObjectURL(blobUrl); capture() }
    iframe.onerror = () => { URL.revokeObjectURL(blobUrl); reject(new Error('幻灯片加载失败')) }
  })
}

// ── PPTX Export (primary) ─────────────────────────────────────────────────

async function exportPPTX(slides, indices, scale) {
  if (typeof PptxGenJS === 'undefined') {
    throw new Error('pptxgenjs 未加载，请检查网络连接后重试')
  }

  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_16x9'   // 10 × 5.625 inches
  pptx.title = 'Slide X Export'

  const total = indices.length

  for (let i = 0; i < indices.length; i++) {
    if (exportCancelled) return
    updateProgress(i, total, `渲染第 ${indices[i] + 1} 页...`)

    const canvas = await renderSlideToCanvas(slides[indices[i]].content, scale)
    const imgData = canvas.toDataURL('image/png')   // base64 png

    const slide = pptx.addSlide()
    // 10 × 5.625 inches = 16:9, matches LAYOUT_16x9
    slide.addImage({
      data: imgData,
      x: 0, y: 0,
      w: 10, h: 5.625
    })
  }

  if (exportCancelled) { cleanupCachedIframe(); return }
  updateProgress(total, total, '正在写入 PPTX 文件...')

  await pptx.writeFile({ fileName: 'presentation.pptx' })
  cleanupCachedIframe()
}

// ── PDF Export ────────────────────────────────────────────────────────────

async function exportPDF(slides, indices, scale) {
  const { jsPDF } = window.jspdf
  const pdf = new jsPDF({
    orientation: 'landscape',
    unit: 'px',
    format: [1280, 720],
    compress: true
  })

  for (let i = 0; i < indices.length; i++) {
    if (exportCancelled) return
    updateProgress(i, indices.length, `渲染第 ${indices[i] + 1} 页...`)

    const canvas = await renderSlideToCanvas(slides[indices[i]].content, scale)
    const imgData = canvas.toDataURL('image/jpeg', 0.92)

    if (i > 0) pdf.addPage([1280, 720], 'landscape')
    pdf.addImage(imgData, 'JPEG', 0, 0, 1280, 720, undefined, 'FAST')
  }

  if (exportCancelled) { cleanupCachedIframe(); return }
  updateProgress(indices.length, indices.length, '正在保存 PDF...')
  pdf.save('presentation.pdf')
  cleanupCachedIframe()
}

// ── Image Export ──────────────────────────────────────────────────────────

async function exportImages(slides, indices, scale, format) {
  const images = []

  for (let i = 0; i < indices.length; i++) {
    if (exportCancelled) return
    updateProgress(i, indices.length, `渲染第 ${indices[i] + 1} 页...`)

    const canvas = await renderSlideToCanvas(slides[indices[i]].content, scale)
    const mimeType = format === 'jpeg' ? 'image/jpeg' : 'image/png'
    images.push({ dataUrl: canvas.toDataURL(mimeType, 0.92), index: indices[i] })
  }

  if (exportCancelled) { cleanupCachedIframe(); return }
  updateProgress(indices.length, indices.length, '打包中...')

  cleanupCachedIframe()
  const ext = format === 'jpeg' ? 'jpg' : 'png'
  if (images.length === 1) {
    downloadDataUrl(images[0].dataUrl, `slide-${images[0].index + 1}.${ext}`)
  } else {
    const zip = new JSZip()
    for (const { dataUrl, index } of images) {
      zip.file(`slide-${String(index + 1).padStart(2, '0')}.${ext}`, dataUrl.split(',')[1], { base64: true })
    }
    const blob = await zip.generateAsync({ type: 'blob' })
    downloadBlob(blob, 'slides-export.zip')
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────

function updateProgress(current, total, text) {
  const bar = document.getElementById('export-progress-bar')
  const label = document.getElementById('export-progress-text')
  const pct = total > 0 ? Math.round((current / total) * 100) : 0
  if (bar) bar.style.width = pct + '%'
  if (label) label.textContent = text || `${current} / ${total}`
}

function parseRange(str, total) {
  const indices = new Set()
  str.split(',').map(s => s.trim()).filter(Boolean).forEach(part => {
    if (part.includes('-')) {
      const [a, b] = part.split('-').map(s => parseInt(s.trim(), 10))
      if (!isNaN(a) && !isNaN(b)) {
        for (let i = Math.max(1, a); i <= Math.min(total, b); i++) indices.add(i - 1)
      }
    } else {
      const n = parseInt(part, 10)
      if (!isNaN(n) && n >= 1 && n <= total) indices.add(n - 1)
    }
  })
  return [...indices].sort((a, b) => a - b)
}

function downloadDataUrl(dataUrl, filename) {
  const a = document.createElement('a')
  a.href = dataUrl
  a.download = filename
  a.click()
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 5000)
}

// ── Editable PPTX Export (Enhanced) ───────────────────────────────────────

// Constants for coordinate conversion
const SLIDE_W = 1280
const SLIDE_H = 720
const PPT_W = 10       // inches (16:9 layout)
const PPT_H = 5.625    // inches

// ── Export Warnings System ────────────────────────────────────────────────

function escapeHtmlForExport(str) {
  if (!str) return ''
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

class ExportWarnings {
  constructor() {
    this.textRenderFailed = []
    this.textScaled = []
    this.imageAddFailed = []
    this.shapeFailed = []
    this.otherWarnings = []
  }

  addTextFailed(text, reason) {
    this.textRenderFailed.push({
      text: text.slice(0, 50) + (text.length > 50 ? '...' : ''),
      reason
    })
  }

  addTextScaled(text, originalSize, newSize) {
    this.textScaled.push({
      text: text.slice(0, 30) + (text.length > 30 ? '...' : ''),
      originalSize,
      newSize
    })
  }

  addImageFailed(src, reason) {
    this.imageAddFailed.push({
      src: src.slice(0, 50),
      reason
    })
  }

  addShapeFailed(type, reason) {
    this.shapeFailed.push({ type, reason })
  }

  addOther(message) {
    this.otherWarnings.push(message)
  }

  hasWarnings() {
    return this.textRenderFailed.length > 0 ||
           this.textScaled.length > 0 ||
           this.imageAddFailed.length > 0 ||
           this.shapeFailed.length > 0 ||
           this.otherWarnings.length > 0
  }

  toSummary() {
    const summary = []
    if (this.textRenderFailed.length > 0) {
      summary.push(`⚠️ ${this.textRenderFailed.length} 个文本元素提取失败`)
    }
    if (this.textScaled.length > 0) {
      summary.push(`📐 ${this.textScaled.length} 个文本已自动缩小字号`)
    }
    if (this.imageAddFailed.length > 0) {
      summary.push(`🖼️ ${this.imageAddFailed.length} 张图片添加失败`)
    }
    if (this.shapeFailed.length > 0) {
      summary.push(`🔷 ${this.shapeFailed.length} 个形状添加失败`)
    }
    if (this.otherWarnings.length > 0) {
      summary.push(`ℹ️ ${this.otherWarnings.length} 条其他警告`)
    }
    return summary
  }

  toDetailedHTML() {
    let html = '<div style="text-align:left;font-size:12px;max-height:200px;overflow-y:auto;">'

    if (this.textScaled.length > 0) {
      html += '<details><summary style="cursor:pointer;margin:6px 0;">📐 字号调整详情</summary><ul>'
      for (const item of this.textScaled.slice(0, 10)) {
        html += `<li>"${escapeHtmlForExport(item.text)}" - ${item.originalSize}pt → ${item.newSize}pt</li>`
      }
      if (this.textScaled.length > 10) {
        html += `<li>...还有 ${this.textScaled.length - 10} 条</li>`
      }
      html += '</ul></details>'
    }

    if (this.textRenderFailed.length > 0) {
      html += '<details><summary style="cursor:pointer;margin:6px 0;">⚠️ 文本提取失败</summary><ul>'
      for (const item of this.textRenderFailed.slice(0, 5)) {
        html += `<li>"${escapeHtmlForExport(item.text)}" - ${escapeHtmlForExport(item.reason)}</li>`
      }
      html += '</ul></details>'
    }

    if (this.imageAddFailed.length > 0) {
      html += '<details><summary style="cursor:pointer;margin:6px 0;">🖼️ 图片添加失败</summary><ul>'
      for (const item of this.imageAddFailed.slice(0, 5)) {
        html += `<li>${escapeHtmlForExport(item.reason)}</li>`
      }
      html += '</ul></details>'
    }

    html += '</div>'
    return html
  }
}

function rgbToHex(rgb) {
  if (!rgb || rgb === 'transparent' || rgb === 'rgba(0, 0, 0, 0)') return null
  const match = rgb.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/)
  if (!match) return null
  const a = match[4] !== undefined ? parseFloat(match[4]) : 1
  if (a === 0) return null  // fully transparent
  const r = parseInt(match[1]), g = parseInt(match[2]), b = parseInt(match[3])
  return ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0').toUpperCase()
}

function mapAlign(cssAlign) {
  const map = { left: 'left', center: 'center', right: 'right', justify: 'justify', start: 'left', end: 'right' }
  return map[cssAlign] || 'left'
}

/**
 * Parse CSS gradient to PptxGenJS gradient format
 * Input length is limited to prevent ReDoS attacks
 */
function parseGradient(bgImage) {
  if (!bgImage || bgImage === 'none') return null

  // Limit input length to prevent ReDoS
  const MAX_GRADIENT_LENGTH = 500
  if (bgImage.length > MAX_GRADIENT_LENGTH) {
    console.warn('Gradient string too long, skipping parse')
    return null
  }

  // Linear gradient: linear-gradient(135deg, #color1, #color2)
  const linearMatch = bgImage.match(/linear-gradient\(\s*([\d.]+)deg\s*,\s*(.+)\s*\)/)
  if (linearMatch) {
    const angle = parseFloat(linearMatch[1])
    const colorStops = linearMatch[2].split(/,(?![^(]*\))/).map(s => s.trim())

    const colors = []
    for (const stop of colorStops) {
      const hex = extractColorFromStop(stop)
      if (hex) colors.push({ color: hex })
    }

    if (colors.length >= 2) {
      return {
        type: 'linear',
        angle: angle,
        stops: colors
      }
    }
  }

  // Radial gradient: radial-gradient(circle, #color1, #color2)
  const radialMatch = bgImage.match(/radial-gradient\(\s*(?:circle|ellipse)?\s*(?:at\s+[\w\s%]+)?\s*,?\s*(.+)\s*\)/)
  if (radialMatch) {
    const colorStops = radialMatch[1].split(/,(?![^(]*\))/).map(s => s.trim())
    const colors = []
    for (const stop of colorStops) {
      const hex = extractColorFromStop(stop)
      if (hex) colors.push({ color: hex })
    }
    if (colors.length >= 2) {
      return {
        type: 'radial',
        stops: colors
      }
    }
  }

  return null
}

function extractColorFromStop(stopStr) {
  // Handle hex colors: #RRGGBB or #RGB
  const hexMatch = stopStr.match(/#([0-9a-fA-F]{6}|[0-9a-fA-F]{3})/)
  if (hexMatch) {
    let hex = hexMatch[1]
    if (hex.length === 3) {
      hex = hex.split('').map(c => c + c).join('')
    }
    return hex.toUpperCase()
  }

  // Handle rgb/rgba
  const rgbMatch = stopStr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/)
  if (rgbMatch) {
    return ((parseInt(rgbMatch[1]) << 16) | (parseInt(rgbMatch[2]) << 8) | parseInt(rgbMatch[3]))
      .toString(16).padStart(6, '0').toUpperCase()
  }

  return null
}

/**
 * Walk a DOM node tree and collect rich text runs compatible with pptxgenjs.
 *
 * Inspired by OpenMAIC lib/export/use-export-pptx.ts formatHTML().
 * Works directly on the live DOM (no AST required) since we have the iframe.
 *
 * @param {Node} node
 * @param {Array} runs  — accumulated array of { text, options }
 * @param {Object} parentStyle  — inherited style object
 * @param {Document} iframeDoc
 */
function walkNodeForRuns(node, runs, parentStyle, iframeDoc) {
  if (node.nodeType === 3 /* TEXT_NODE */) {
    const text = node.textContent
    if (text) runs.push({ text, options: { ...parentStyle } })
    return
  }
  if (node.nodeType !== 1 /* ELEMENT_NODE */) return

  const tag = node.tagName.toLowerCase()
  if (tag === 'script' || tag === 'style') return

  const cs = iframeDoc.defaultView.getComputedStyle(node)
  if (cs.display === 'none' || cs.visibility === 'hidden') return

  // Build style for this node, inheriting from parent
  const style = { ...parentStyle }

  // Semantic tag overrides
  if (tag === 'b' || tag === 'strong') style.bold = true
  if (tag === 'i' || tag === 'em')     style.italic = true
  if (tag === 'u')                     style.underline = true

  // Computed style overrides (more specific than tag semantics)
  const fw = parseInt(cs.fontWeight)
  if (!isNaN(fw)) style.bold = fw >= 600
  if (cs.fontStyle === 'italic' || cs.fontStyle === 'oblique') style.italic = true
  const textDec = cs.textDecorationLine || cs.textDecoration || ''
  if (textDec.includes('underline')) style.underline = true

  const color = rgbToHex(cs.color)
  if (color) style.color = color

  const fsPx = parseFloat(cs.fontSize)
  if (fsPx > 0) style.fontSize = Math.round(fsPx * 72 / 96)

  const ff = extractFontFamily(cs.fontFamily)
  if (ff) style.fontFace = ff

  // Block-level tags inject a line break before their content
  const isBlock = ['div', 'p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'br', 'tr'].includes(tag)
  if (isBlock && runs.length > 0) {
    runs[runs.length - 1].options.breakLine = true
  }

  for (const child of node.childNodes) {
    walkNodeForRuns(child, runs, style, iframeDoc)
  }
}

/**
 * Extract rich text runs from an element for use with pptxgenjs addText().
 * Returns null when the element has no meaningful inline formatting variation
 * (plain text is sufficient in that case).
 *
 * @param {Element} el
 * @param {Document} iframeDoc
 * @param {CSSStyleDeclaration} elStyle - already-computed style for el (avoids duplicate call)
 * @returns {Array | null}
 */
function extractRichTextRuns(el, iframeDoc, elStyle) {
  // Leaf nodes have no child elements — no formatting variation is possible
  if (el.children.length === 0) return null

  const baseStyle = {
    color:     rgbToHex(elStyle.color) || '333333',
    bold:      parseInt(elStyle.fontWeight) >= 600,
    italic:    (elStyle.fontStyle === 'italic' || elStyle.fontStyle === 'oblique'),
    underline: (elStyle.textDecorationLine || elStyle.textDecoration || '').includes('underline'),
    fontSize:  Math.round(parseFloat(elStyle.fontSize) * 72 / 96),
    fontFace:  extractFontFamily(elStyle.fontFamily),
  }

  const runs = []
  walkNodeForRuns(el, runs, baseStyle, iframeDoc)

  // Filter newline-only runs but preserve breakLine markers and space word-separators.
  // Whitespace text nodes between block elements carry breakLine from walkNodeForRuns;
  // discarding them with a naive trim() filter silently drops paragraph breaks.
  // Space-only runs (e.g. between <b>Hello</b> and <i>World</i>) must be kept.
  const nonEmpty = []
  for (const run of runs) {
    if (run.text.replace(/[\n\r]/g, '').length > 0) {
      nonEmpty.push(run)
    } else if (run.options.breakLine && nonEmpty.length > 0) {
      // Transfer the breakLine to the previous kept run so it isn't lost
      nonEmpty[nonEmpty.length - 1].options.breakLine = true
    }
  }
  if (nonEmpty.length === 0) return null

  // Only return as rich runs when there IS formatting variation between runs
  // (avoids overhead for plain single-style elements)
  const hasVariation = nonEmpty.some(r =>
    r.options.bold      !== baseStyle.bold      ||
    r.options.italic    !== baseStyle.italic    ||
    r.options.underline !== baseStyle.underline ||
    r.options.color     !== baseStyle.color     ||
    r.options.fontSize  !== baseStyle.fontSize  ||
    r.options.fontFace  !== baseStyle.fontFace
  )
  return hasVariation ? nonEmpty : null
}

/**
 * Extract text elements with improved styling
 */
function extractTextElements(iframeDoc, warnings = null) {
  const elements = []
  const addedTexts = new Set()

  // Collect elements with data-role first (semantic priority), then fallback selectors
  const semanticEls = Array.from(iframeDoc.querySelectorAll('[data-role]'))
  const semanticSet = new Set(semanticEls)

  const fallbackEls = Array.from(iframeDoc.querySelectorAll(
    'h1,h2,h3,h4,h5,h6,p,li,td,th,span,div,' +
    '[class*="title"],[class*="heading"],[class*="subtitle"],' +
    '[class*="content"],[class*="text"],[class*="body"],' +
    '[class*="label"],[class*="stat"],[class*="number"]'
  ))

  // Semantic elements come first; fallback elements fill in the rest
  const candidates = [...semanticEls, ...fallbackEls.filter(el => !semanticSet.has(el))]

  for (const el of candidates) {
    const role = el.getAttribute('data-role')

    // Skip notes and chart elements — notes are hidden by design, charts handled separately
    if (role === 'notes' || role === 'chart') continue

    // Skip text inside chart containers (already captured as images)
    if (!role && el.closest('[data-role="chart"]')) continue

    const text = (el.innerText || el.textContent || '').trim()
    if (!text || text.length < 1) continue

    // Semantic elements bypass dedup (same text may appear in title + body legitimately)
    if (!role && addedTexts.has(text)) continue

    // For non-semantic elements, skip containers that have text children
    if (!role) {
      const hasTextChildren = Array.from(el.children).some(child =>
        (child.innerText || '').trim().length > 0
      )
      if (hasTextChildren && el.children.length > 0) continue
    }

    const style = iframeDoc.defaultView.getComputedStyle(el)
    if (style.display === 'none' || style.visibility === 'hidden') continue
    if (parseFloat(style.opacity) < 0.1) continue

    const rect = el.getBoundingClientRect()
    if (rect.width < 5 || rect.height < 5) continue
    if (rect.left > SLIDE_W || rect.top > SLIDE_H) continue

    const baseFontSizePt = Math.round(parseFloat(style.fontSize) * 72 / 96)
    const fontWeight = parseInt(style.fontWeight)
    const fontFace = extractFontFamily(style.fontFamily)

    const widthPt = (rect.width / 96) * 72
    const heightPt = (rect.height / 96) * 72
    const optimizedFontSize = calculateFontSize(widthPt, heightPt, text, baseFontSizePt, fontFace)

    if (warnings && optimizedFontSize < baseFontSizePt * 0.7) {
      warnings.addTextScaled(text, baseFontSizePt, optimizedFontSize)
    }

    // Attempt rich text extraction to preserve inline formatting (bold/italic/color/size variation)
    // Pass the already-computed style to avoid a redundant getComputedStyle() call
    const richRuns = extractRichTextRuns(el, iframeDoc, style)
    // Scale ratio: if calculateFontSize shrunk the base size, proportionally shrink run sizes too
    const scaleRatio = baseFontSizePt > 0 ? optimizedFontSize / baseFontSizePt : 1

    elements.push({
      type: 'text',
      role: role || null,
      text,
      richRuns,   // Array<{text, options}> or null (null = use plain text fallback)
      scaleRatio,
      x: Math.max(0, rect.left / SLIDE_W * PPT_W),
      y: Math.max(0, rect.top / SLIDE_H * PPT_H),
      w: Math.min(Math.max(0.2, rect.width / SLIDE_W * PPT_W), PPT_W),
      h: Math.min(Math.max(0.2, rect.height / SLIDE_H * PPT_H), PPT_H),
      fontSize: optimizedFontSize,
      bold: fontWeight >= 600,
      italic: style.fontStyle.includes('italic'),
      underline: style.textDecoration.includes('underline'),
      color: rgbToHex(style.color) || '333333',
      align: mapAlign(style.textAlign),
      fontFace,
    })
    addedTexts.add(text)
  }

  return elements
}

/**
 * Calculate optimal font size to prevent text overflow in PowerPoint.
 * Uses canvas measureText() for precise per-character width measurement.
 * Supports CJK character-level line wrapping.
 */
function calculateFontSize(widthPt, heightPt, text, baseFontSize, fontFamily = 'Arial') {
  if (!text || widthPt <= 0 || heightPt <= 0 || baseFontSize <= 0) {
    return Math.max(8, Math.min(baseFontSize || 12, 96))
  }

  const ctx = _getMeasureCanvas().getContext('2d')
  const toPx = pt => pt * 96 / 72
  const widthPx = toPx(widthPt)
  const heightPx = toPx(heightPt)

  let fontSize = Math.min(baseFontSize, 96)
  const MIN_SIZE = 8

  while (fontSize >= MIN_SIZE) {
    ctx.font = `${toPx(fontSize)}px ${fontFamily}`
    const lineHeightPx = toPx(fontSize) * 1.3

    // Character-level wrapping (works for CJK and Latin)
    let line = ''
    let lines = 1
    for (const char of text) {
      const testLine = line + char
      if (ctx.measureText(testLine).width > widthPx && line.length > 0) {
        lines++
        line = char
      } else {
        line = testLine
      }
    }

    if (lines * lineHeightPx <= heightPx) return fontSize

    fontSize -= fontSize <= 22 ? 1 : 2
  }

  return MIN_SIZE
}

/** Singleton offscreen canvas for text measurement (avoids repeated DOM creation) */
let _measureCanvas = null
function _getMeasureCanvas() {
  if (!_measureCanvas) {
    _measureCanvas = document.createElement('canvas')
    _measureCanvas.width = 2000
    _measureCanvas.height = 100
  }
  return _measureCanvas
}

function extractFontFamily(cssFontFamily) {
  // Return first valid font family
  const fonts = cssFontFamily.split(',').map(f => f.trim().replace(/['"]/g, ''))
  // Map system fonts to PowerPoint-friendly alternatives
  const fontMap = {
    '-apple-system': 'Arial',
    'BlinkMacSystemFont': 'Arial',
    'Segoe UI': 'Segoe UI',
    'system-ui': 'Arial',
    'sans-serif': 'Arial',
    'serif': 'Times New Roman',
    'monospace': 'Courier New',
  }
  for (const font of fonts) {
    if (fontMap[font]) return fontMap[font]
    if (!font.includes('system') && !font.includes('-apple')) return font
  }
  return 'Arial'
}

/**
 * Extract shape elements (rectangles, circles, etc.)
 */
function extractShapeElements(iframeDoc) {
  const shapes = []
  const candidates = iframeDoc.querySelectorAll(
    '[class*="shape"],[class*="box"],[class*="card"],[class*="block"],' +
    '[class*="icon"],[class*="badge"],[class*="circle"],[class*="rect"]'
  )

  for (const el of candidates) {
    const style = iframeDoc.defaultView.getComputedStyle(el)
    if (style.display === 'none' || style.visibility === 'hidden') continue

    const rect = el.getBoundingClientRect()
    if (rect.width < 10 || rect.height < 10) continue
    if (rect.left > SLIDE_W || rect.top > SLIDE_H) continue

    const bgColor = rgbToHex(style.backgroundColor)
    const borderRadius = parseFloat(style.borderRadius) || 0
    const borderColor = rgbToHex(style.borderColor)
    const borderWidth = parseFloat(style.borderWidth) || 0

    // Skip elements that are likely text containers
    if (!bgColor && !borderColor) continue

    // Determine shape type based on border-radius
    const isCircle = borderRadius >= Math.min(rect.width, rect.height) / 2

    shapes.push({
      type: 'shape',
      shapeType: isCircle ? 'ellipse' : 'rect',
      x: Math.max(0, rect.left / SLIDE_W * PPT_W),
      y: Math.max(0, rect.top / SLIDE_H * PPT_H),
      w: Math.max(0.1, rect.width / SLIDE_W * PPT_W),
      h: Math.max(0.1, rect.height / SLIDE_H * PPT_H),
      fill: bgColor ? { color: bgColor } : null,
      line: borderWidth > 0 && borderColor ? { color: borderColor, pt: borderWidth } : null,
      rectRadius: !isCircle && borderRadius > 0 ? borderRadius / 100 : 0,
    })
  }

  return shapes
}

/**
 * Extract slide background (solid color or gradient)
 */
function extractSlideBackground(iframeDoc) {
  const candidates = [
    iframeDoc.querySelector('section[data-slide]'),
    iframeDoc.querySelector('section'),
    iframeDoc.body,
    iframeDoc.querySelector('[class*="slide"]'),
  ]

  for (const el of candidates) {
    if (!el) continue
    const style = iframeDoc.defaultView.getComputedStyle(el)

    // Check for gradient first
    const gradient = parseGradient(style.backgroundImage)
    if (gradient) {
      return { type: 'gradient', gradient }
    }

    // Fall back to solid color
    const hex = rgbToHex(style.backgroundColor)
    if (hex) {
      return { type: 'solid', color: hex }
    }
  }

  return { type: 'solid', color: 'FFFFFF' }
}

/**
 * Extract images and convert to base64
 */
async function extractImageElements(iframeDoc, warnings = null) {
  const images = []
  const imgElements = iframeDoc.querySelectorAll('img')

  for (const img of imgElements) {
    if (!img.src) continue

    const rect = img.getBoundingClientRect()
    if (rect.width < 10 || rect.height < 10) continue
    if (rect.left > SLIDE_W || rect.top > SLIDE_H) continue

    try {
      // Convert image to base64
      const canvas = document.createElement('canvas')
      canvas.width = img.naturalWidth || rect.width
      canvas.height = img.naturalHeight || rect.height
      const ctx = canvas.getContext('2d')
      ctx.drawImage(img, 0, 0)
      const dataUrl = canvas.toDataURL('image/png')

      images.push({
        type: 'image',
        data: dataUrl,
        x: Math.max(0, rect.left / SLIDE_W * PPT_W),
        y: Math.max(0, rect.top / SLIDE_H * PPT_H),
        w: Math.max(0.1, rect.width / SLIDE_W * PPT_W),
        h: Math.max(0.1, rect.height / SLIDE_H * PPT_H),
      })
    } catch (e) {
      // Skip images that fail to convert (CORS issues, etc.)
      console.warn('Failed to extract image:', e)
      if (warnings) {
        warnings.addImageFailed(img.src, e.message || 'CORS or canvas error')
      }
    }
  }

  return images
}

/**
 * Extract canvas/SVG/chart elements as images for editable PPTX export.
 * Handles: data-role="chart", <canvas>, <svg>, and common chart library containers.
 */
async function extractChartElements(iframeDoc, warnings = null) {
  const charts = []
  const selector = [
    '[data-role="chart"]',
    'canvas',
    'svg:not([aria-hidden="true"])',
    '[class*="chart"]',
    '[class*="echarts"]',
    '[class*="highcharts"]',
    '[class*="plotly"]',
  ].join(',')

  const matched = Array.from(iframeDoc.querySelectorAll(selector))

  for (const el of matched) {
    // Skip elements nested inside another matched element
    // (e.g. a <canvas> inside a [data-role="chart"] container — capture only the outermost)
    if (matched.some(other => other !== el && other.contains(el))) continue
    const rect = el.getBoundingClientRect()
    if (rect.width < 10 || rect.height < 10) continue
    if (rect.left > SLIDE_W || rect.top > SLIDE_H) continue

    try {
      let dataUrl = null
      const tag = el.tagName.toLowerCase()

      if (tag === 'canvas') {
        // Direct canvas → PNG
        dataUrl = el.toDataURL('image/png')
      } else if (tag === 'svg') {
        // Serialize SVG → blob URL → canvas → PNG
        const svgData = new XMLSerializer().serializeToString(el)
        const svgBlob = new Blob([svgData], { type: 'image/svg+xml' })
        const url = URL.createObjectURL(svgBlob)
        dataUrl = await new Promise((res, rej) => {
          const img = new Image()
          img.onload = () => {
            const c = document.createElement('canvas')
            c.width = rect.width
            c.height = rect.height
            c.getContext('2d').drawImage(img, 0, 0, rect.width, rect.height)
            URL.revokeObjectURL(url)
            res(c.toDataURL('image/png'))
          }
          img.onerror = (e) => { URL.revokeObjectURL(url); rej(e) }
          img.src = url
        })
      } else {
        // Generic element — use html2canvas on the iframe body, crop to element rect
        if (typeof html2canvas === 'undefined') continue
        const fullCanvas = await html2canvas(el, { useCORS: true, scale: 1, logging: false })
        dataUrl = fullCanvas.toDataURL('image/png')
      }

      if (!dataUrl || dataUrl === 'data:,') continue

      charts.push({
        type: 'chart-image',
        data: dataUrl,
        x: Math.max(0, rect.left / SLIDE_W * PPT_W),
        y: Math.max(0, rect.top / SLIDE_H * PPT_H),
        w: Math.max(0.1, rect.width / SLIDE_W * PPT_W),
        h: Math.max(0.1, rect.height / SLIDE_H * PPT_H),
      })
    } catch (e) {
      console.warn('Failed to extract chart element:', e)
      if (warnings) warnings.addImageFailed('chart element', e.message || 'capture error')
    }
  }

  return charts
}

function loadSlideForExtraction(htmlContent, warnings = null) {
  return new Promise((resolve, reject) => {
    const iframe = getOrCreateIframe()

    const blob = new Blob([htmlContent], { type: 'text/html' })
    const blobUrl = URL.createObjectURL(blob)
    iframe.src = blobUrl
    iframe.onload = async () => {
      URL.revokeObjectURL(blobUrl)
      // Wait for fonts and styles to load
      await new Promise(r => setTimeout(r, 400))
      try {
        const doc = iframe.contentDocument
        const textElements = extractTextElements(doc, warnings)
        const shapeElements = extractShapeElements(doc)
        const imageElements = await extractImageElements(doc, warnings)
        const chartElements = await extractChartElements(doc, warnings)
        const background = extractSlideBackground(doc)
        resolve({ textElements, shapeElements, imageElements, chartElements, background })
      } catch (err) {
        reject(err)
      }
    }
    iframe.onerror = (e) => { URL.revokeObjectURL(blobUrl); reject(e) }
  })
}

async function exportEditablePPTX(slides, indices) {
  if (typeof PptxGenJS === 'undefined') {
    throw new Error('pptxgenjs 未加载，请检查网络连接后重试')
  }

  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_16x9'
  pptx.title = 'Slide X Export'
  pptx.author = 'Slide X'

  const total = indices.length
  const warnings = new ExportWarnings()

  for (let i = 0; i < indices.length; i++) {
    if (exportCancelled) return
    updateProgress(i, total, `提取第 ${indices[i] + 1} 页元素...`)

    const { textElements, shapeElements, imageElements, chartElements, background } =
      await loadSlideForExtraction(slides[indices[i]].content, warnings)

    const slide = pptx.addSlide()

    // Apply background
    if (background.type === 'gradient' && background.gradient) {
      // PptxGenJS gradient support
      if (background.gradient.type === 'linear' && background.gradient.stops.length >= 2) {
        slide.background = {
          color: background.gradient.stops[0].color,
          // PptxGenJS has limited gradient support, use first color as fallback
        }
      } else {
        slide.background = { color: background.gradient.stops[0]?.color || 'FFFFFF' }
      }
    } else {
      slide.background = { color: background.color || 'FFFFFF' }
    }

    // Add shapes first (background layer)
    for (const shape of shapeElements) {
      try {
        const shapeOpts = {
          x: shape.x,
          y: shape.y,
          w: shape.w,
          h: shape.h,
        }
        if (shape.fill) shapeOpts.fill = shape.fill
        if (shape.line) shapeOpts.line = shape.line
        if (shape.rectRadius) shapeOpts.rectRadius = shape.rectRadius

        slide.addShape(shape.shapeType, shapeOpts)
      } catch (e) {
        console.warn('Failed to add shape:', e)
        warnings.addShapeFailed(shape.shapeType, e.message)
      }
    }

    // Add images
    for (const img of imageElements) {
      try {
        slide.addImage({
          data: img.data,
          x: img.x,
          y: img.y,
          w: img.w,
          h: img.h,
        })
      } catch (e) {
        console.warn('Failed to add image:', e)
        warnings.addImageFailed('slide image', e.message)
      }
    }

    // Add chart/canvas/SVG elements as images
    for (const chart of chartElements) {
      try {
        slide.addImage({
          data: chart.data,
          x: chart.x,
          y: chart.y,
          w: chart.w,
          h: chart.h,
        })
      } catch (e) {
        console.warn('Failed to add chart image:', e)
        warnings.addImageFailed('chart element', e.message)
      }
    }

    // Add text elements (top layer)
    for (const el of textElements) {
      try {
        const textOpts = {
          x: el.x,
          y: el.y,
          w: el.w,
          h: el.h,
          fontSize: el.fontSize,
          bold: el.bold,
          italic: el.italic,
          underline: el.underline ? { style: 'sng' } : undefined,
          color: el.color,
          align: el.align,
          valign: 'top',
          fontFace: el.fontFace,
          wrap: true,
        }

        if (el.richRuns && el.richRuns.length > 0) {
          // Rich text: scale each run's fontSize by the same ratio used for the container
          const scaledRuns = el.richRuns.map(run => ({
            text: run.text,
            options: {
              ...run.options,
              fontSize: run.options.fontSize
                ? Math.max(8, Math.round(run.options.fontSize * el.scaleRatio))
                : el.fontSize,
              underline: run.options.underline ? { style: 'sng' } : undefined,
            },
          }))
          slide.addText(scaledRuns, textOpts)
        } else {
          slide.addText(el.text, textOpts)
        }
      } catch (e) {
        console.warn('Failed to add text:', e)
        warnings.addTextFailed(el.text, e.message)
      }
    }

    // Add speaker notes
    const notes = slides[indices[i]].notes
    if (notes) {
      try { slide.addNotes(notes) } catch (e) { /* pptxgenjs version may not support addNotes */ }
    }
  }

  if (exportCancelled) { cleanupCachedIframe(); return }
  updateProgress(total, total, '正在写入可编辑 PPTX 文件...')
  await pptx.writeFile({ fileName: 'presentation-editable.pptx' })
  cleanupCachedIframe()

  // Show warnings if any
  if (warnings.hasWarnings()) {
    showExportWarnings(warnings)
  }
}

/**
 * Display export warnings to user
 */
function showExportWarnings(warnings) {
  const summary = warnings.toSummary()
  const details = warnings.toDetailedHTML()

  // Create a simple warning modal
  const existingModal = document.getElementById('export-warnings-modal')
  if (existingModal) existingModal.remove()

  const modal = document.createElement('div')
  modal.id = 'export-warnings-modal'
  modal.className = 'modal-backdrop'
  modal.innerHTML = `
    <div class="modal" style="width:380px;">
      <div class="modal-header">
        导出完成
        <button class="modal-close" onclick="this.closest('.modal-backdrop').remove()">×</button>
      </div>
      <div class="modal-body">
        <div style="margin-bottom:12px;color:var(--text-secondary);">
          文件已导出，但有以下提示：
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;">
          ${summary.map(s => `<div style="font-size:13px;">${escapeHtmlForExport(s)}</div>`).join('')}
        </div>
        ${details}
      </div>
      <div class="modal-footer">
        <button class="btn-primary" onclick="this.closest('.modal-backdrop').remove()">
          确定
        </button>
      </div>
    </div>
  `

  document.body.appendChild(modal)
}
