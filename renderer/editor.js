/**
 * editor.js — CodeMirror 6 source editor + visual editing
 */

let cmView = null
let cmInitialized = false
let pendingContent = null
let debounceTimer = null
let applyCallback = null
let firstEditCallback = null   // fired on the FIRST user keystroke per edit session
let firstEditFired = false     // reset when content is set programmatically
let isSetting = false          // suppresses auto-apply during programmatic content set

const DEBOUNCE_MS = 600

// ── Initialization ────────────────────────────────────────────────────────

export async function initEditor(container, onApply, onFirstEdit) {
  applyCallback = onApply
  firstEditCallback = onFirstEdit || null

  try {
    const { EditorView, basicSetup } = await import('https://esm.sh/codemirror@6?bundle')
    const { EditorState } = await import('https://esm.sh/@codemirror/state@6?bundle')
    const { keymap } = await import('https://esm.sh/@codemirror/view@6?bundle')
    const { indentWithTab, defaultKeymap } = await import('https://esm.sh/@codemirror/commands@6?bundle')
    const { html: htmlLang } = await import('https://esm.sh/@codemirror/lang-html@6?bundle')

    const startState = EditorState.create({
      doc: pendingContent || '',
      extensions: [
        basicSetup,
        htmlLang(),
        keymap.of([
          ...defaultKeymap,
          indentWithTab,
          { key: 'Ctrl-s', run: () => { triggerApply(); return true; }, preventDefault: true },
          { key: 'Cmd-s', run: () => { triggerApply(); return true; }, preventDefault: true },
        ]),
        EditorView.updateListener.of(update => {
          if (update.docChanged && !isSetting) {
            // First real keystroke in this edit session → push undo state immediately
            if (!firstEditFired && firstEditCallback) {
              firstEditFired = true
              firstEditCallback()
            }
            scheduleAutoApply()
          }
        }),
        EditorView.theme({
          '&': { height: '100%', fontSize: '12px' },
          '.cm-scroller': { overflow: 'auto', fontFamily: "'JetBrains Mono', 'Fira Code', 'Consolas', monospace" }
        })
      ]
    })

    cmView = new EditorView({
      state: startState,
      parent: container
    })

    cmInitialized = true

    if (pendingContent !== null) {
      setEditorContent(pendingContent)
      pendingContent = null
    }
  } catch (err) {
    console.error('CodeMirror init failed:', err)
    // Fallback: use a textarea
    const ta = document.createElement('textarea')
    ta.style.cssText = 'width:100%;height:100%;background:#16162a;color:#e8e8f0;font-family:monospace;font-size:12px;padding:12px;border:none;outline:none;resize:none;'
    ta.value = pendingContent || ''
    container.appendChild(ta)
    ta.addEventListener('input', () => {
      if (!firstEditFired && firstEditCallback) {
        firstEditFired = true
        firstEditCallback()
      }
      scheduleAutoApply()
    })
    ta.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault()
        triggerApply()
      }
    })
    // Shim
    cmView = {
      _ta: ta,
      state: { doc: { toString: () => ta.value } },
      dispatch: () => {},
      destroy: () => ta.remove()
    }
    cmInitialized = true
    pendingContent = null
  }
}

// ── Content ────────────────────────────────────────────────────────────────

export function setEditorContent(html) {
  if (!cmInitialized || !cmView) {
    pendingContent = html
    return
  }

  if (cmView._ta) {
    firstEditFired = false
    cmView._ta.value = html
    return
  }

  // Suppress auto-apply while we programmatically set content
  isSetting = true
  firstEditFired = false  // next real keystroke is a fresh edit session
  if (debounceTimer) { clearTimeout(debounceTimer); debounceTimer = null }

  cmView.dispatch({
    changes: { from: 0, to: cmView.state.doc.length, insert: html }
  })

  // Re-enable after microtask flush so the updateListener has fired
  setTimeout(() => { isSetting = false }, 50)
}

export function getEditorContent() {
  if (!cmView) return ''
  if (cmView._ta) return cmView._ta.value
  return cmView.state.doc.toString()
}

export function focusEditor() {
  if (cmView && !cmView._ta) cmView.focus()
  else if (cmView?._ta) cmView._ta.focus()
}

// ── Auto Apply ────────────────────────────────────────────────────────────

function scheduleAutoApply() {
  if (debounceTimer) clearTimeout(debounceTimer)
  debounceTimer = setTimeout(() => {
    triggerApply()
  }, DEBOUNCE_MS)
}

function triggerApply() {
  if (debounceTimer) { clearTimeout(debounceTimer); debounceTimer = null; }
  if (applyCallback) applyCallback(getEditorContent())
}

export function applyEditorChanges() {
  triggerApply()
}

export function setApplyCallback(fn) {
  applyCallback = fn
}

// ── Visual Editor ──────────────────────────────────────────────────────────

let visualEditActive = false
let currentEditIframe = null
let visualChangeCallback = null

export function enableVisualEdit(iframe, onChange) {
  if (visualEditActive) return
  currentEditIframe = iframe
  visualChangeCallback = onChange
  visualEditActive = true

  const iframeDoc = iframe.contentDocument
  if (!iframeDoc) return

  // Make text elements editable
  const selectors = 'h1, h2, h3, h4, h5, h6, p, span, li, td, th, div, a, button, label'
  const elements = iframeDoc.querySelectorAll(selectors)

  elements.forEach(el => {
    // Skip elements that only contain other block elements
    if (!hasDirectText(el)) return
    el.setAttribute('data-ve-original', el.innerHTML)
    el.style.cursor = 'text'
    el.style.outline = 'none'

    el.addEventListener('click', handleVisualClick, true)
    el.addEventListener('dblclick', handleVisualDblClick, true)
  })

  // Click outside to deactivate editing
  iframeDoc.addEventListener('click', (e) => {
    const editing = iframeDoc.querySelector('[contenteditable="true"]')
    if (editing && !editing.contains(e.target)) {
      commitEdit(editing)
    }
  }, true)
}

export function disableVisualEdit(iframe) {
  if (!iframe || !iframe.contentDocument) return
  visualEditActive = false
  currentEditIframe = null

  const iframeDoc = iframe.contentDocument
  const elements = iframeDoc.querySelectorAll('[data-ve-original]')
  elements.forEach(el => {
    el.style.cursor = ''
    el.style.outline = ''
    el.removeAttribute('contenteditable')
    el.removeEventListener('click', handleVisualClick, true)
    el.removeEventListener('dblclick', handleVisualDblClick, true)
    el.removeAttribute('data-ve-original')
  })
}

function handleVisualClick(e) {
  // Single click: show hover highlight
  if (!this.getAttribute('contenteditable')) {
    this.style.outline = '2px solid rgba(230, 126, 34, 0.4)'
  }
}

function handleVisualDblClick(e) {
  e.stopPropagation()
  e.preventDefault()
  const el = this

  // Deactivate any current editing
  const doc = el.ownerDocument
  const prev = doc.querySelector('[contenteditable="true"]')
  if (prev && prev !== el) commitEdit(prev)

  // Activate editing
  el.contentEditable = 'true'
  el.style.outline = '2px solid #E67E22'
  el.style.background = 'rgba(230, 126, 34, 0.08)'

  // Focus is required on Windows — without it the element is contentEditable
  // but never receives keyboard input (macOS handles this implicitly on dblclick)
  el.focus()

  // Place cursor at end
  const range = doc.createRange()
  const sel = doc.defaultView?.getSelection() ?? doc.getSelection()
  if (sel) {
    range.selectNodeContents(el)
    range.collapse(false)
    sel.removeAllRanges()
    sel.addRange(range)
  }

  // Clean up any previous listeners to prevent duplicates
  const handleBlur = () => {
    commitEdit(el)
    el.removeEventListener('keydown', handleKeydown)
  }
  const handleKeydown = (e) => {
    if (e.key === 'Escape') { el.blur() }
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); el.blur() }
  }

  el.addEventListener('blur', handleBlur, { once: true })
  el.addEventListener('keydown', handleKeydown)
}

function commitEdit(el) {
  el.contentEditable = 'false'
  el.style.outline = ''
  el.style.background = ''

  const original = el.getAttribute('data-ve-original')
  if (el.innerHTML !== original) {
    el.setAttribute('data-ve-original', el.innerHTML)
    if (visualChangeCallback && currentEditIframe) {
      visualChangeCallback(currentEditIframe.contentDocument.documentElement.outerHTML)
    }
  }
}

function hasDirectText(el) {
  for (const node of el.childNodes) {
    if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) return true
  }
  return false
}
