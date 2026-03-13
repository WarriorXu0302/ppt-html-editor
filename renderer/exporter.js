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
    iframe.onerror = () => { reject(new Error('幻灯片加载失败')) }
  })
}

// ── PPTX Export (primary) ─────────────────────────────────────────────────

async function exportPPTX(slides, indices, scale) {
  if (typeof PptxGenJS === 'undefined') {
    throw new Error('pptxgenjs 未加载，请检查网络连接后重试')
  }

  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_16x9'   // 10 × 5.625 inches
  pptx.title = 'PPT Editor Export'

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
  if (images.length === 1) {
    const ext = format === 'jpeg' ? 'jpg' : 'png'
    downloadDataUrl(images[0].dataUrl, `slide-${images[0].index + 1}.${ext}`)
  } else {
    const zip = new JSZip()
    for (const { dataUrl, index } of images) {
      const ext = format === 'jpeg' ? 'jpg' : 'png'
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
 */
function parseGradient(bgImage) {
  if (!bgImage || bgImage === 'none') return null

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
 * Extract text elements with improved styling
 */
function extractTextElements(iframeDoc, warnings = null) {
  const elements = []
  const addedTexts = new Set()

  const candidates = iframeDoc.querySelectorAll(
    'h1,h2,h3,h4,h5,h6,p,li,td,th,span,div,' +
    '[class*="title"],[class*="heading"],[class*="subtitle"],' +
    '[class*="content"],[class*="text"],[class*="body"],' +
    '[class*="label"],[class*="stat"],[class*="number"]'
  )

  for (const el of candidates) {
    const text = (el.innerText || el.textContent || '').trim()
    if (!text || text.length < 1) continue
    if (addedTexts.has(text)) continue

    // Check if this is a leaf text node (no child elements with their own text)
    const hasTextChildren = Array.from(el.children).some(child =>
      (child.innerText || '').trim().length > 0
    )
    if (hasTextChildren && el.children.length > 0) continue

    const style = iframeDoc.defaultView.getComputedStyle(el)
    if (style.display === 'none' || style.visibility === 'hidden') continue
    if (parseFloat(style.opacity) < 0.1) continue

    const rect = el.getBoundingClientRect()
    if (rect.width < 5 || rect.height < 5) continue
    if (rect.left > SLIDE_W || rect.top > SLIDE_H) continue

    const baseFontSizePt = Math.round(parseFloat(style.fontSize) * 72 / 96)
    const fontWeight = parseInt(style.fontWeight)

    // Calculate optimal font size to prevent overflow
    const widthPt = (rect.width / 96) * 72
    const heightPt = (rect.height / 96) * 72
    const optimizedFontSize = calculateFontSize(widthPt, heightPt, text, baseFontSizePt)

    // Track if font was scaled down significantly
    if (warnings && optimizedFontSize < baseFontSizePt * 0.7) {
      warnings.addTextScaled(text, baseFontSizePt, optimizedFontSize)
    }

    elements.push({
      type: 'text',
      text,
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
      fontFace: extractFontFamily(style.fontFamily),
    })
    addedTexts.add(text)
  }

  return elements
}

/**
 * Calculate optimal font size to prevent text overflow in PowerPoint
 * CJK-aware: Chinese/Japanese/Korean characters are wider than Latin
 */
function calculateFontSize(widthPt, heightPt, text, baseFontSize) {
  // Guard against invalid inputs
  if (!text || widthPt <= 0 || heightPt <= 0 || baseFontSize <= 0) {
    return Math.max(8, Math.min(baseFontSize || 12, 96))
  }

  // Count CJK and non-CJK characters
  const cjkCount = [...text].filter(c =>
    (c >= '\u4e00' && c <= '\u9fff') ||  // CJK Unified Ideographs
    (c >= '\u3040' && c <= '\u30ff') ||  // Hiragana + Katakana
    (c >= '\uac00' && c <= '\ud7af')     // Korean Hangul
  ).length
  const nonCjkCount = text.length - cjkCount

  // Estimate text width: CJK chars ~1em, Latin chars ~0.5em on average
  const estimatedTextWidth = (cjkCount * 1.0 + nonCjkCount * 0.55) * baseFontSize

  // Calculate lines needed
  const linesNeeded = Math.ceil(estimatedTextWidth / Math.max(widthPt, 1))
  const lineHeight = baseFontSize * 1.3  // typical line height

  // Calculate total height required
  const totalHeight = linesNeeded * lineHeight

  // If text would overflow, scale down
  if (totalHeight > heightPt && heightPt > 0) {
    const scaleFactor = heightPt / totalHeight
    const scaledSize = Math.floor(baseFontSize * scaleFactor)
    return Math.max(8, Math.min(scaledSize, 96))  // clamp between 8-96pt
  }

  return Math.max(8, Math.min(baseFontSize, 96))
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
        const background = extractSlideBackground(doc)
        resolve({ textElements, shapeElements, imageElements, background })
      } catch (err) {
        reject(err)
      }
    }
    iframe.onerror = (e) => { reject(e) }
  })
}

async function exportEditablePPTX(slides, indices) {
  if (typeof PptxGenJS === 'undefined') {
    throw new Error('pptxgenjs 未加载，请检查网络连接后重试')
  }

  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_16x9'
  pptx.title = 'PPT Editor Export'
  pptx.author = 'PPT HTML Editor'

  const total = indices.length
  const warnings = new ExportWarnings()

  for (let i = 0; i < indices.length; i++) {
    if (exportCancelled) return
    updateProgress(i, total, `提取第 ${indices[i] + 1} 页元素...`)

    const { textElements, shapeElements, imageElements, background } =
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
        slide.addText(el.text, textOpts)
      } catch (e) {
        console.warn('Failed to add text:', e)
        warnings.addTextFailed(el.text, e.message)
      }
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
          ${summary.map(s => `<div style="font-size:13px;">${s}</div>`).join('')}
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
