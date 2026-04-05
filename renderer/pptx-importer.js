/**
 * pptx-importer.js — PPTX → HTML slide converter
 *
 * Reads a .pptx file (ArrayBuffer) via JSZip + DOMParser and produces
 * an array of slide objects compatible with app.js state.
 *
 * Format: returns [{ content, rawHtml, title, notes }]
 * content = standalone full HTML document (1280×720px) for iframe rendering
 */

// EMU (English Metric Unit) constants
const SLIDE_W_EMU = 9144000   // 10 inches (standard 16:9 widescreen)
const SLIDE_H_EMU = 5143500   // ~5.625 inches
const SLIDE_W_PX  = 1280
const SLIDE_H_PX  = 720

/**
 * Import a .pptx file from a byte array (from Electron IPC)
 * @param {number[]} byteArray  – plain Array of bytes (from main.js)
 * @returns {Promise<Array>}    – slide objects
 */
export async function importPPTX(byteArray) {
  if (typeof JSZip === 'undefined') {
    throw new Error('JSZip 未加载，请检查网络连接后重试')
  }

  const buffer = new Uint8Array(byteArray).buffer
  const zip = await JSZip.loadAsync(buffer)

  // Determine slide order from presentation relationships
  const slideOrder = await parseSlideOrder(zip)
  if (!slideOrder.length) throw new Error('未找到幻灯片内容')

  // Cache slide layouts by path — most presentations reuse the same 1-2 layouts
  const layoutCache = new Map()

  const results = await Promise.all(slideOrder.map(async (slideName, i) => {
    const [slideXml, slideRelsXml] = await Promise.all([
      zip.file(`ppt/slides/${slideName}.xml`)?.async('text'),
      zip.file(`ppt/slides/_rels/${slideName}.xml.rels`)?.async('text')
    ])
    if (!slideXml) return null

    // Parse rels once; pass parsed doc to both resolveSlideLayout and convertSlideToHtml
    const relsDoc = slideRelsXml
      ? new DOMParser().parseFromString(slideRelsXml, 'application/xml')
      : null
    if (relsDoc?.querySelector('parsererror')) {
      console.warn('Invalid XML in slide relationships:', slideName)
      return null
    }
    const slideLayoutXml = await resolveSlideLayout(zip, relsDoc, layoutCache)

    const { html, title, notes } = await convertSlideToHtml(slideXml, relsDoc, slideLayoutXml, zip)
    return { content: html, rawHtml: html, title: title || `幻灯片 ${i + 1}`, notes }
  }))

  return results.filter(Boolean)
}

// ── Slide Order ──────────────────────────────────────────────────────────

async function parseSlideOrder(zip) {
  const relsXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('text')
  if (!relsXml) return []

  const doc = new DOMParser().parseFromString(relsXml, 'application/xml')
  const rels = Array.from(doc.querySelectorAll('Relationship'))
    .filter(r => r.getAttribute('Type')?.endsWith('/slide'))
    .sort((a, b) => {
      // Sort by numeric id to preserve order
      const idA = parseInt(a.getAttribute('Id')?.replace(/\D/g, '') || '0')
      const idB = parseInt(b.getAttribute('Id')?.replace(/\D/g, '') || '0')
      return idA - idB
    })

  return rels.map(r => {
    const target = r.getAttribute('Target') || ''
    // e.g. "slides/slide1.xml" → "slide1"
    return target.replace(/^.*\//, '').replace(/\.xml$/, '')
  }).filter(Boolean)
}

// ── Slide Layout ─────────────────────────────────────────────────────────

async function resolveSlideLayout(zip, relsDoc, layoutCache) {
  if (!relsDoc) return null
  const layoutRel = Array.from(relsDoc.querySelectorAll('Relationship'))
    .find(r => r.getAttribute('Type')?.endsWith('/slideLayout'))
  if (!layoutRel) return null

  const target = layoutRel.getAttribute('Target') || ''
  // Target is relative to ppt/slides/, so we need ppt/slideLayouts/...
  const layoutPath = 'ppt/' + target.replace(/^\.\.\//, '')
  if (layoutCache.has(layoutPath)) return layoutCache.get(layoutPath)
  const xml = await zip.file(layoutPath)?.async('text') || null
  layoutCache.set(layoutPath, xml)
  return xml
}

// ── Main Converter ────────────────────────────────────────────────────────

async function convertSlideToHtml(slideXml, relsDoc, slideLayoutXml, zip) {
  const doc = new DOMParser().parseFromString(slideXml, 'application/xml')
  if (doc.querySelector('parsererror')) {
    throw new Error('Invalid slide XML: ' + doc.querySelector('parsererror').textContent.slice(0, 200))
  }
  const layoutDoc = slideLayoutXml
    ? new DOMParser().parseFromString(slideLayoutXml, 'application/xml')
    : null

  // Extract background color
  const bgColor = extractBackground(doc, layoutDoc)

  // Extract text boxes
  const textBoxes = extractTextBoxes(doc)

  // Extract images
  const images = await extractImages(doc, relsDoc, zip)

  // Extract notes
  const notes = extractNotes(doc)

  // Find title
  const titleBox = textBoxes.find(t => t.isTitle)
  const title = titleBox?.text || ''

  // Build HTML
  const html = buildSlideHtml(bgColor, textBoxes, images)

  return { html, title, notes }
}

// ── Background ────────────────────────────────────────────────────────────

function extractBackground(doc, layoutDoc) {
  // Try slide background first, then layout
  for (const d of [doc, layoutDoc].filter(Boolean)) {
    const solidFill = d.querySelector('bg solidFill srgbClr, background solidFill srgbClr')
    if (solidFill) {
      return '#' + (solidFill.getAttribute('val') || 'FFFFFF')
    }
    // Gradient fallback — just use first stop color
    const gradStop = d.querySelector('bg gradFill gsLst gs:first-child srgbClr')
    if (gradStop) return '#' + (gradStop.getAttribute('val') || 'FFFFFF')
  }
  return '#FFFFFF'
}

// ── Text Boxes ────────────────────────────────────────────────────────────

function extractTextBoxes(doc) {
  const boxes = []
  const MAX_TEXT_BOXES_PER_SLIDE = 100

  for (const sp of doc.querySelectorAll('sp')) {
    if (boxes.length >= MAX_TEXT_BOXES_PER_SLIDE) break
    const phType = sp.querySelector('ph')?.getAttribute('type') || ''
    const isTitle = phType === 'title' || phType === 'ctrTitle'

    // Position & size
    const xfrm = sp.querySelector('spPr xfrm') || sp.querySelector('xfrm')
    if (!xfrm) continue

    const off = xfrm.querySelector('off')
    const ext = xfrm.querySelector('ext')
    if (!off || !ext) continue

    const x = emuToPx(parseInt(off.getAttribute('x') || 0), SLIDE_W_EMU, SLIDE_W_PX)
    const y = emuToPx(parseInt(off.getAttribute('y') || 0), SLIDE_H_EMU, SLIDE_H_PX)
    const w = emuToPx(parseInt(ext.getAttribute('cx') || 0), SLIDE_W_EMU, SLIDE_W_PX)
    const h = emuToPx(parseInt(ext.getAttribute('cy') || 0), SLIDE_H_EMU, SLIDE_H_PX)

    if (w < 5 || h < 5) continue

    // Extract text with paragraph runs
    const paragraphs = []
    for (const para of sp.querySelectorAll('p')) {
      const runs = []
      let paraText = ''
      for (const r of para.querySelectorAll('r')) {
        const t = r.querySelector('t')?.textContent || ''
        if (!t) continue
        const rPr = r.querySelector('rPr')
        const sz = parseInt(rPr?.getAttribute('sz') || '0') / 100 || null
        const bold = rPr?.getAttribute('b') === '1'
        const color = rPr?.querySelector('solidFill srgbClr')?.getAttribute('val') || null
        runs.push({ text: t, sz, bold, color })
        paraText += t
      }
      if (paraText) paragraphs.push({ text: paraText, runs })
    }

    if (!paragraphs.length) continue

    const fullText = paragraphs.map(p => p.text).join('\n')
    boxes.push({ x, y, w, h, isTitle, paragraphs, text: fullText })
  }

  return boxes
}

// ── Images ────────────────────────────────────────────────────────────────

async function extractImages(doc, relsDoc, zip) {
  const images = []
  if (!relsDoc) return images

  const MAX_IMAGES_PER_SLIDE = 50

  const relMap = {}
  for (const r of relsDoc.querySelectorAll('Relationship')) {
    relMap[r.getAttribute('Id')] = r.getAttribute('Target')
  }

  for (const pic of doc.querySelectorAll('pic')) {
    if (images.length >= MAX_IMAGES_PER_SLIDE) break
    const rEmbed = pic.querySelector('blipFill blip')?.getAttribute('r:embed') ||
                   pic.querySelector('blipFill blip')?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed')
    if (!rEmbed) continue

    const target = relMap[rEmbed]
    if (!target) continue

    const imgPath = 'ppt/' + target.replace(/^\.\.\//, '')
    const imgData = await zip.file(imgPath)?.async('base64')
    if (!imgData) continue

    const ext = imgPath.split('.').pop().toLowerCase()
    const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg'
               : ext === 'gif' ? 'image/gif'
               : ext === 'svg' ? 'image/svg+xml'
               : 'image/png'

    const xfrm = pic.querySelector('spPr xfrm') || pic.querySelector('xfrm')
    const off = xfrm?.querySelector('off')
    const exts = xfrm?.querySelector('ext')
    if (!off || !exts) continue

    images.push({
      src: `data:${mime};base64,${imgData}`,
      x: emuToPx(parseInt(off.getAttribute('x') || 0), SLIDE_W_EMU, SLIDE_W_PX),
      y: emuToPx(parseInt(off.getAttribute('y') || 0), SLIDE_H_EMU, SLIDE_H_PX),
      w: emuToPx(parseInt(exts.getAttribute('cx') || 0), SLIDE_W_EMU, SLIDE_W_PX),
      h: emuToPx(parseInt(exts.getAttribute('cy') || 0), SLIDE_H_EMU, SLIDE_H_PX),
    })
  }

  return images
}

// ── Notes ─────────────────────────────────────────────────────────────────

function extractNotes(_doc) {
  // PPTX speaker notes live in ppt/notesSlides/ (separate files), not in slide XML.
  // Extraction from notes slide files is not yet implemented.
  return ''
}

// ── HTML Builder ──────────────────────────────────────────────────────────

function buildSlideHtml(bgColor, textBoxes, images) {
  const imgHtml = images.map(img => `
  <img src="${escapeAttr(img.src)}" style="
    position:absolute;
    left:${img.x}px; top:${img.y}px;
    width:${img.w}px; height:${img.h}px;
    object-fit:contain;
  " />`).join('')

  const textHtml = textBoxes.map(box => {
    const paraHtml = box.paragraphs.map(p => {
      const runHtml = p.runs.map(r => {
        const style = [
          r.sz ? `font-size:${r.sz}px` : '',
          r.bold ? 'font-weight:bold' : '',
          r.color ? `color:#${r.color}` : '',
        ].filter(Boolean).join(';')
        return style ? `<span style="${escapeAttr(style)}">${escapeHtml(r.text)}</span>` : escapeHtml(r.text)
      }).join('')
      return `<p style="margin:0;line-height:1.3;">${runHtml}</p>`
    }).join('')

    const role = box.isTitle ? 'title' : 'body'
    return `
  <div data-role="${role}" style="
    position:absolute;
    left:${box.x}px; top:${box.y}px;
    width:${box.w}px; height:${box.h}px;
    overflow:hidden;
    box-sizing:border-box;
    padding:4px;
  ">${paraHtml}</div>`
  }).join('')

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0; padding: 0; overflow: hidden;
    width: ${SLIDE_W_PX}px; height: ${SLIDE_H_PX}px;
    background: ${bgColor};
    font-family: 'PingFang SC', 'Microsoft YaHei', Arial, sans-serif;
  }
</style>
</head>
<body>${imgHtml}${textHtml}
</body>
</html>`
}

// ── Helpers ───────────────────────────────────────────────────────────────

function emuToPx(emu, totalEmu, totalPx) {
  return Math.round((emu / totalEmu) * totalPx)
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function escapeAttr(str) {
  return String(str).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}
