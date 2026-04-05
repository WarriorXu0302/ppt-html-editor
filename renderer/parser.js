/**
 * parser.js — Multi-strategy HTML PPT parser
 * Returns: { format, docHead, docOuter, slides: [{ content, title, rawHtml }] }
 *
 * content = standalone full HTML document for rendering in an iframe
 * rawHtml  = the "editable unit" shown in source editor
 */

export function parseHTML(htmlString) {
  const parser = new DOMParser()
  const doc = parser.parseFromString(htmlString, 'text/html')
  const headHTML = doc.head.innerHTML

  // ── Strategy 1: section[data-slide] ──────────────────────────────────────
  const sectionSlides = doc.querySelectorAll('section[data-slide]')
  if (sectionSlides.length > 0) {
    return {
      format: 'section-data-slide',
      docHead: headHTML,
      docOuter: null,
      slides: Array.from(sectionSlides).map((section, i) => ({
        content: buildStandaloneDoc(headHTML, section.outerHTML),
        rawHtml: section.outerHTML,
        title: section.getAttribute('data-title') || `幻灯片 ${i + 1}`,
        notes: extractNotesFromElement(section)
      }))
    }
  }

  // ── Strategy 2: iframe[srcdoc] ────────────────────────────────────────────
  const iframeSlides = doc.querySelectorAll('iframe[srcdoc]')
  if (iframeSlides.length > 0) {
    // Preserve the outer wrapper HTML for reconstruction
    return {
      format: 'iframe-srcdoc',
      docHead: headHTML,
      docOuter: buildOuterTemplate(doc),
      slides: Array.from(iframeSlides).map((iframe, i) => {
        const srcdocHTML = iframe.getAttribute('srcdoc')
        // Use DOM parsing instead of regex for title extraction (safer)
        const rawTitle = extractTitleFromHTML(srcdocHTML) || `幻灯片 ${i + 1}`
        return {
          content: srcdocHTML,
          rawHtml: srcdocHTML,
          title: rawTitle,
          notes: extractNotesFromHTML(srcdocHTML)
        }
      })
    }
  }

  // ── Strategy 3: section / div.slide / div[class*="slide"] ─────────────────
  const genericSlides = doc.querySelectorAll('section, div.slide, div[class*="slide"]')
  if (genericSlides.length > 1) {
    return {
      format: 'generic',
      docHead: headHTML,
      docOuter: null,
      slides: Array.from(genericSlides).map((el, i) => ({
        content: buildStandaloneDoc(headHTML, el.outerHTML),
        rawHtml: el.outerHTML,
        title: `幻灯片 ${i + 1}`,
        notes: extractNotesFromElement(el)
      }))
    }
  }

  // ── Strategy 4: whole document as single slide ───────────────────────────
  return {
    format: 'raw',
    docHead: headHTML,
    docOuter: null,
    slides: [{
      content: htmlString,
      rawHtml: htmlString,
      title: doc.title || '幻灯片 1',
      notes: extractNotesFromHTML(htmlString)
    }]
  }
}

/**
 * Reconstruct the full saveable document from slides
 */
export function reconstructHTML(format, docHead, docOuter, slides) {
  switch (format) {
    case 'section-data-slide': {
      // Extract section outerHTML from each slide's standalone doc
      const sections = slides.map((slide, i) => {
        const doc = new DOMParser().parseFromString(slide.content, 'text/html')
        const section = doc.querySelector('section[data-slide]')
        if (section) return section.outerHTML
        // Fallback: wrap body content
        return `<section data-slide="${i + 1}" data-title="${escapeAttr(slide.title)}">${doc.body.innerHTML}</section>`
      })
      return `<!DOCTYPE html>
<html>
<head>
${docHead}
</head>
<body>
${sections.join('\n')}
</body>
</html>`
    }

    case 'iframe-srcdoc': {
      const items = slides.map(slide => {
        const encoded = escapeForAttr(slide.content)
        return `    <div class="ppt_page_iframe_content">
      <iframe class="ppt_page_iframe" srcdoc="${encoded}"></iframe>
    </div>`
      })
      return docOuter
        ? docOuter.replace('{{SLIDES}}', items.join('\n'))
        : buildIframeSrcdocWrapper(items.join('\n'))
    }

    case 'generic': {
      const elements = slides.map((slide, i) => {
        const doc = new DOMParser().parseFromString(slide.content, 'text/html')
        return doc.body.firstElementChild
          ? doc.body.firstElementChild.outerHTML
          : `<div class="slide">${doc.body.innerHTML}</div>`
      })
      return `<!DOCTYPE html>
<html>
<head>
${docHead}
</head>
<body>
${elements.join('\n')}
</body>
</html>`
    }

    case 'raw':
    default:
      return slides[0]?.content || ''
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────

function buildStandaloneDoc(headHTML, bodyContent) {
  return `<!DOCTYPE html>
<html>
<head>
${headHTML}
<style>
  body { margin: 0; padding: 0; overflow: hidden; }
  section[data-slide] {
    display: block !important;
    width: 1280px;
    height: 720px;
    position: relative;
    overflow: hidden;
  }
</style>
</head>
<body>
${bodyContent}
</body>
</html>`
}

function buildOuterTemplate(doc) {
  // Try to capture the original wrapper structure with a placeholder for slides
  const parent = doc.querySelector('.ppt_iframe_parent')
  if (!parent) return null
  // Build wrapper HTML
  const headHTML = doc.head.innerHTML
  return `<!DOCTYPE html>
<html lang="en">
<head>
${headHTML}
</head>
<body>
  <div class="ppt_iframe_parent">
{{SLIDES}}
  </div>
</body>
</html>`
}

function buildIframeSrcdocWrapper(slidesHTML) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    * { box-sizing: border-box; }
    .ppt_iframe_parent {
      max-width: 1280px;
      width: 100%;
      height: 100vh;
      margin: 0 auto;
      scroll-snap-type: y mandatory;
      overflow-y: auto;
    }
    .ppt_page_iframe_content {
      width: 100%;
      aspect-ratio: 16 / 9;
      overflow: hidden;
      border-radius: 12px;
      scroll-snap-align: start;
      margin-bottom: 24px;
    }
    .ppt_page_iframe {
      display: block;
      border: none;
      flex-shrink: 0;
      overflow: hidden;
      width: 1280px;
      height: 720px;
      transform: scale(var(--ppt-html-zoom, 1));
      transform-origin: left top;
    }
  </style>
  <script>
    function updateElementZoom() {
      const el = document.querySelector('.ppt_iframe_parent');
      if (el) el.style.setProperty('--ppt-html-zoom', '' + el.clientWidth / 1280);
    }
    window.addEventListener('DOMContentLoaded', updateElementZoom);
    window.addEventListener('resize', updateElementZoom);
  <\/script>
</head>
<body>
  <div class="ppt_iframe_parent">
${slidesHTML}
  </div>
</body>
</html>`
}

function escapeForAttr(html) {
  return html
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function escapeAttr(str) {
  return (str || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

/**
 * Extract speaker notes text from a DOM element (data-role="notes")
 */
function extractNotesFromElement(el) {
  const notesEl = el.querySelector('[data-role="notes"]')
  return notesEl ? (notesEl.textContent || '').trim() : ''
}

/**
 * Extract speaker notes text from an HTML string
 */
function extractNotesFromHTML(html) {
  if (!html) return ''
  try {
    const tempDoc = new DOMParser().parseFromString(html, 'text/html')
    return extractNotesFromElement(tempDoc.body)
  } catch {
    return ''
  }
}

/**
 * Extract title from HTML string using DOM parsing (safer than regex)
 */
function extractTitleFromHTML(html) {
  if (!html) return null
  try {
    const tempDoc = new DOMParser().parseFromString(html, 'text/html')
    const titleEl = tempDoc.querySelector('title')
    return titleEl ? titleEl.textContent.trim() : null
  } catch {
    return null
  }
}
