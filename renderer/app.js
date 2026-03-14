/**
 * app.js — Main controller: state management, event coordination
 */

import { parseHTML, reconstructHTML } from './parser.js'
import { importPPTX } from './pptx-importer.js'
import { initEditor, setEditorContent, getEditorContent, applyEditorChanges,
         enableVisualEdit, disableVisualEdit } from './editor.js'
import { initExporter, showExportModal } from './exporter.js'
import { initAIPanel } from './ai-panel.js'

// ── State ──────────────────────────────────────────────────────────────────

const state = {
  slides: [],         // [{ content: string, rawHtml: string, title: string, notes: string }]
  currentIndex: 0,
  filePath: null,
  isDirty: false,
  format: 'unknown',
  docHead: '',
  docOuter: null,
  undoStacks: {},     // { index: [html, ...] }
  redoStacks: {},     // { index: [html, ...] }
  editMode: 'visual', // 'visual' | 'source'
  previewMode: false, // true = interactive preview, false = editable
  sidebarOpen: true,
  aiPanelOpen: false,
  editorOpen: false,
  isResizingEditor: false,
  editorHeight: 280,
  thumbObserver: null,
  visualEditActive: false
}

// Expose state globally for exporter
window.appState = state

// ── Init ───────────────────────────────────────────────────────────────────

export async function initApp() {
  // Platform class for Mac toolbar spacing
  if (window.electronAPI.platform === 'darwin') {
    document.body.classList.add('mac')
  }

  // Initialize sub-modules
  await initEditor(
    document.getElementById('editor-container'),
    // apply: don't push undo here — first-edit callback already did it
    (html) => applySlideContent(html, false),
    // first-edit: push undo state the moment the user starts typing
    () => {
      if (state.slides.length > 0) pushUndoState(state.currentIndex)
    }
  )

  initExporter()

  await initAIPanel({
    onGenerate: (html) => loadHTMLContent(html, null),
    onModify: (html) => replaceCurrentSlide(html),
    getAppState: () => state
  })

  setupToolbar()
  setupKeyboard()
  setupDragDrop()
  setupEditorResize()
  setupMenuListeners()
  setupEditorModeListeners()
  setupNotesPanel()

  // Show welcome
  showWelcome(true)
}

// ── Toolbar ────────────────────────────────────────────────────────────────

function setupToolbar() {
  // Sidebar toggle
  document.getElementById('sidebar-toggle-btn').addEventListener('click', toggleSidebar)

  // File ops
  document.getElementById('open-btn').addEventListener('click', openFile)
  document.getElementById('import-pptx-btn').addEventListener('click', importPPTXFile)
  document.getElementById('save-btn').addEventListener('click', saveFile)

  // Undo / Redo
  document.getElementById('undo-btn').addEventListener('click', undo)
  document.getElementById('redo-btn').addEventListener('click', redo)

  // AI Panel toggle
  document.getElementById('ai-toggle-btn').addEventListener('click', toggleAIPanel)

  // Edit mode toggle
  document.getElementById('edit-mode-btn').addEventListener('click', toggleEditMode)

  // Presentation
  document.getElementById('present-btn').addEventListener('click', openPresentation)

  // Export
  document.getElementById('export-btn').addEventListener('click', () => showExportModal(state.slides.length))

  // Settings
  // (handled in ai-panel.js, but we add the gear button here)

  // Nav buttons
  document.getElementById('prev-btn').addEventListener('click', () => navigate(-1))
  document.getElementById('next-btn').addEventListener('click', () => navigate(1))

  // Preview/Edit mode toggle
  document.getElementById('preview-toggle-btn').addEventListener('click', togglePreviewMode)
}

function togglePreviewMode() {
  state.previewMode = !state.previewMode

  const btn = document.getElementById('preview-toggle-btn')
  const icon = btn.querySelector('.toggle-icon')
  const label = btn.querySelector('.toggle-label')

  if (state.previewMode) {
    // Preview mode: disable visual editing
    btn.classList.add('preview-mode')
    icon.textContent = '👁️'
    label.textContent = '预览'

    const iframe = document.getElementById('main-preview')
    if (state.visualEditActive) {
      disableVisualEdit(iframe)
      state.visualEditActive = false
    }
  } else {
    // Edit mode: enable visual editing
    btn.classList.remove('preview-mode')
    icon.textContent = '✏️'
    label.textContent = '编辑'

    if (state.editMode === 'visual' && state.slides.length > 0) {
      const iframe = document.getElementById('main-preview')
      activateVisualEdit(iframe)
    }
  }
}

function setupEditorModeListeners() {
  document.getElementById('editor-apply-btn').addEventListener('click', () => applyEditorChanges())

  // Editor close button
  const editorCloseBtn = document.createElement('button')
  editorCloseBtn.className = 'editor-action-btn'
  editorCloseBtn.textContent = '✕'
  editorCloseBtn.title = '关闭编辑器'
  editorCloseBtn.addEventListener('click', () => closeEditor())
  document.getElementById('editor-toolbar').appendChild(editorCloseBtn)
}

// ── Menu Listeners ────────────────────────────────────────────────────────

function setupMenuListeners() {
  window.electronAPI.onMenuEvent('menu-open', () => openFile())
  window.electronAPI.onMenuEvent('menu-save', () => saveFile())
  window.electronAPI.onMenuEvent('menu-save-as', () => saveFileAs())
  window.electronAPI.onMenuEvent('menu-undo', () => undo())
  window.electronAPI.onMenuEvent('menu-redo', () => redo())
  window.electronAPI.onMenuEvent('menu-new', () => newFile())
  window.electronAPI.onMenuEvent('menu-open-file', (filePath) => openFileByPath(filePath))

  document.addEventListener('app:save', () => saveFile())
}

// ── Keyboard ──────────────────────────────────────────────────────────────

function setupKeyboard() {
  document.addEventListener('keydown', (e) => {
    const isEditorFocused = document.getElementById('editor-drawer').classList.contains('open') &&
      document.activeElement?.closest('#editor-container')

    if (isEditorFocused) return // Let CodeMirror handle

    const ctrl = e.ctrlKey || e.metaKey

    if (ctrl && e.key === 'z') { e.preventDefault(); undo(); }
    else if (ctrl && (e.key === 'y' || (e.shiftKey && e.key === 'z'))) { e.preventDefault(); redo(); }
    else if (ctrl && e.key === 's' && !e.shiftKey) { e.preventDefault(); saveFile(); }
    else if (ctrl && e.key === 's' && e.shiftKey) { e.preventDefault(); saveFileAs(); }
    else if (ctrl && e.key === 'o') { e.preventDefault(); openFile(); }
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
      if (!isInputFocused()) { e.preventDefault(); navigate(-1); }
    } else if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
      if (!isInputFocused()) { e.preventDefault(); navigate(1); }
    } else if (e.key === 'F5' || (ctrl && e.key === 'F5')) {
      e.preventDefault(); openPresentation()
    }
  })
}

function isInputFocused() {
  const tag = document.activeElement?.tagName
  return ['INPUT', 'TEXTAREA', 'SELECT'].includes(tag) ||
    document.activeElement?.contentEditable === 'true'
}

// ── Drag & Drop ────────────────────────────────────────────────────────────

function setupDragDrop() {
  document.addEventListener('dragover', (e) => {
    e.preventDefault()
    document.body.classList.add('drag-over')
  })
  document.addEventListener('dragleave', (e) => {
    if (e.relatedTarget === null) document.body.classList.remove('drag-over')
  })
  document.addEventListener('drop', async (e) => {
    e.preventDefault()
    document.body.classList.remove('drag-over')
    const file = e.dataTransfer.files[0]
    if (file && (file.name.endsWith('.html') || file.name.endsWith('.htm'))) {
      await openFileByPath(file.path)
    }
  })
}

// ── Editor Resize ──────────────────────────────────────────────────────────

function setupEditorResize() {
  const handle = document.getElementById('editor-resize-handle')
  const drawer = document.getElementById('editor-drawer')

  handle.addEventListener('mousedown', (e) => {
    e.preventDefault()
    state.isResizingEditor = true
    const startY = e.clientY
    const startH = state.editorHeight

    const onMove = (e) => {
      const delta = startY - e.clientY
      const newH = Math.max(120, Math.min(600, startH + delta))
      state.editorHeight = newH
      drawer.style.height = newH + 'px'
    }

    const onUp = () => {
      state.isResizingEditor = false
      document.removeEventListener('mousemove', onMove)
      document.removeEventListener('mouseup', onUp)
    }

    document.addEventListener('mousemove', onMove)
    document.addEventListener('mouseup', onUp)
  })
}

// ── File Operations ────────────────────────────────────────────────────────

async function openFile() {
  const result = await window.electronAPI.showOpenDialog()
  if (result.canceled || !result.filePaths?.[0]) return

  await openFileByPath(result.filePaths[0])
}

async function openFileByPath(filePath) {
  if (!await checkUnsaved()) return
  try {
    const content = await window.electronAPI.readFile(filePath)
    loadHTMLContent(content, filePath)

    // Add to recent files
    const config = await window.electronAPI.getConfig()
    const recent = [filePath, ...(config.recentFiles || []).filter(f => f !== filePath)].slice(0, 10)
    await window.electronAPI.setConfig({ recentFiles: recent })
  } catch (err) {
    console.error('Failed to open file:', err)
    alert('打开文件失败：' + err.message)
  }
}

async function newFile() {
  if (!await checkUnsaved()) return
  state.slides = []
  state.currentIndex = 0
  state.filePath = null
  state.format = 'unknown'
  state.undoStacks = {}
  state.redoStacks = {}
  setDirty(false)
  renderAll()
  showWelcome(true)
  updateTitle()
}

async function importPPTXFile() {
  if (!await checkUnsaved()) return
  try {
    const result = await window.electronAPI.openPptxFile()
    if (!result) return  // cancelled

    const importedSlides = await importPPTX(result.data)
    if (!importedSlides.length) {
      alert('未能从该 PPTX 文件中提取幻灯片内容')
      return
    }

    state.slides = importedSlides
    state.format = 'section-data-slide'
    state.docHead = ''
    state.docOuter = null
    state.currentIndex = 0
    state.undoStacks = {}
    state.redoStacks = {}
    state.filePath = null
    setDirty(true)
    showWelcome(false)
    renderThumbnails()
    renderPreview()
    updateTitle()
  } catch (e) {
    console.error('PPTX import failed:', e)
    alert('导入失败：' + (e.message || '未知错误'))
  }
}

async function saveFile() {
  if (!state.filePath) return saveFileAs()
  if (state.slides.length === 0) return

  const html = reconstructHTML(state.format, state.docHead, state.docOuter, state.slides)
  try {
    await window.electronAPI.writeFile(state.filePath, html)
    setDirty(false)
    // Notify main process that save is complete
    window.electronAPI.notifySaveComplete()
  } catch (err) {
    // Permission denied → fall back to Save As (e.g. opened from a read-only path)
    if (err.message.includes('EACCES') || err.message.includes('EPERM') ||
        err.message.includes('permission denied') || err.message.includes('operation not permitted')) {
      return saveFileAs()
    }
    alert('保存失败：' + err.message)
  }
}

async function saveFileAs() {
  if (state.slides.length === 0) return

  const defaultPath = state.filePath || 'presentation.html'
  const result = await window.electronAPI.showSaveDialog(defaultPath)
  if (result.canceled || !result.filePath) return

  const html = reconstructHTML(state.format, state.docHead, state.docOuter, state.slides)
  try {
    await window.electronAPI.writeFile(result.filePath, html)
    state.filePath = result.filePath
    setDirty(false)
    updateTitle()
    // Notify main process that save is complete
    window.electronAPI.notifySaveComplete()
  } catch (err) {
    alert('保存失败：' + err.message)
  }
}

async function checkUnsaved() {
  if (!state.isDirty || state.slides.length === 0) return true
  const result = await window.electronAPI.showMessageBox({
    type: 'question',
    buttons: ['保存', '不保存', '取消'],
    defaultId: 0,
    cancelId: 2,
    message: '当前文件有未保存的修改',
    detail: '是否保存？'
  })
  if (result.response === 0) { await saveFile(); return !state.isDirty }
  if (result.response === 1) return true
  return false
}

// ── Load Content ──────────────────────────────────────────────────────────

function loadHTMLContent(htmlString, filePath) {
  const parsed = parseHTML(htmlString)

  state.slides = parsed.slides
  state.format = parsed.format
  state.docHead = parsed.docHead
  state.docOuter = parsed.docOuter
  state.currentIndex = 0
  state.undoStacks = {}
  state.redoStacks = {}
  state.filePath = filePath

  setDirty(filePath === null) // Dirty if generated (no file path)
  showWelcome(false)
  renderAll()
  updateTitle()
}

// ── Rendering ─────────────────────────────────────────────────────────────

function renderAll() {
  renderThumbnails()
  renderPreview()
  updateNavButtons()  // also handles empty-slides case (renderPreview returns early then)
}

// Thumbnail rendering with IntersectionObserver lazy loading

function renderThumbnails() {
  const container = document.getElementById('thumbnails')
  container.innerHTML = ''

  if (state.thumbObserver) {
    state.thumbObserver.disconnect()
    state.thumbObserver = null
  }

  if (state.slides.length === 0) return

  const useLazy = state.slides.length > 20
  const thumbWidth = 176 // sidebar 200px - 24px padding

  if (useLazy) {
    state.thumbObserver = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          const item = entry.target
          const idx = parseInt(item.dataset.index)
          if (!item.querySelector('iframe')) {
            createThumbIframe(item, idx, thumbWidth)
          }
          state.thumbObserver.unobserve(item)
        }
      })
    }, { root: container, rootMargin: '100px', threshold: 0.01 })
  }

  state.slides.forEach((_slide, i) => {
    const item = createThumbItem(i, thumbWidth, !useLazy)
    container.appendChild(item)
    if (useLazy) state.thumbObserver.observe(item)
  })

  updateThumbnailActive()
}

function createThumbItem(index, thumbWidth, renderNow) {
  const item = document.createElement('div')
  item.className = `thumb-item${index === state.currentIndex ? ' active' : ''}`
  item.dataset.index = index

  const thumbH = Math.round(thumbWidth * (720 / 1280))
  item.style.width = thumbWidth + 'px'

  const wrapper = document.createElement('div')
  wrapper.className = 'thumb-wrapper'
  wrapper.style.height = thumbH + 'px'
  wrapper.style.width = thumbWidth + 'px'

  if (renderNow) {
    createThumbIframe(item, index, thumbWidth)
  } else {
    const placeholder = document.createElement('div')
    placeholder.className = 'thumb-placeholder'
    placeholder.textContent = index + 1
    wrapper.appendChild(placeholder)
    item.appendChild(wrapper)

    const badge = document.createElement('div')
    badge.className = 'thumb-badge'
    badge.textContent = index + 1
    item.appendChild(badge)
  }

  item.addEventListener('click', () => goToSlide(index))

  return item
}

function createThumbIframe(item, index, thumbWidth) {
  // Remove placeholder
  item.innerHTML = ''

  const thumbH = Math.round(thumbWidth * (720 / 1280))
  const scale = thumbWidth / 1280

  const wrapper = document.createElement('div')
  wrapper.className = 'thumb-wrapper'
  wrapper.style.cssText = `width:${thumbWidth}px;height:${thumbH}px;`

  const iframe = document.createElement('iframe')
  iframe.style.transform = `scale(${scale})`
  iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin')
  iframe.setAttribute('title', `幻灯片 ${index + 1}`)
  iframe.style.overflow = 'hidden'

  wrapper.appendChild(iframe)
  item.appendChild(wrapper)

  const badge = document.createElement('div')
  badge.className = 'thumb-badge'
  badge.textContent = index + 1
  item.appendChild(badge)

  // Load content via blob URL
  const content = state.slides[index].content
  const blob = new Blob([content], { type: 'text/html' })
  const blobUrl = URL.createObjectURL(blob)
  iframe.src = blobUrl

  // Clean up blob URL after load or on error
  const cleanup = () => URL.revokeObjectURL(blobUrl)
  iframe.onload = cleanup
  iframe.onerror = cleanup
}

function refreshThumb(index) {
  const item = document.querySelector(`.thumb-item[data-index="${index}"]`)
  if (!item) return
  const iframe = item.querySelector('iframe')
  if (!iframe) return

  const content = state.slides[index].content
  const blob = new Blob([content], { type: 'text/html' })
  const blobUrl = URL.createObjectURL(blob)
  iframe.src = blobUrl

  // Clean up blob URL after load or on error
  const cleanup = () => URL.revokeObjectURL(blobUrl)
  iframe.onload = cleanup
  iframe.onerror = cleanup
}

function updateThumbnailActive() {
  document.querySelectorAll('.thumb-item').forEach((item, i) => {
    item.classList.toggle('active', i === state.currentIndex)
  })
  // Scroll active into view
  const active = document.querySelector('.thumb-item.active')
  if (active) active.scrollIntoView({ block: 'nearest', behavior: 'smooth' })
}

// Main preview rendering

function renderPreview() {
  if (state.slides.length === 0) return

  const iframe = document.getElementById('main-preview')

  // Disable old visual editing
  if (state.visualEditActive) {
    disableVisualEdit(iframe)
    state.visualEditActive = false
  }

  const content = state.slides[state.currentIndex].content
  const blob = new Blob([content], { type: 'text/html' })
  const blobUrl = URL.createObjectURL(blob)
  iframe.src = blobUrl
  iframe.onload = () => {
    URL.revokeObjectURL(blobUrl)
    // Only activate visual edit if NOT in preview mode
    if (state.editMode === 'visual' && !state.previewMode) {
      activateVisualEdit(iframe)
    }
    scalePreview()
  }

  updatePageInfo()
  updateNavButtons()
  updateEditorContent()
  updateNotesPanel()
}

function scalePreview() {
  const container = document.getElementById('preview-container')
  const wrapper = document.getElementById('preview-wrapper')
  const iframe = document.getElementById('main-preview')

  const cw = container.clientWidth - 40  // 20px margin each side
  const ch = container.clientHeight - 40

  const scale = Math.min(cw / 1280, ch / 720, 1)

  iframe.style.transform = `scale(${scale})`
  wrapper.style.width = Math.round(1280 * scale) + 'px'
  wrapper.style.height = Math.round(720 * scale) + 'px'

  // Adjust iframe origin to fill the wrapper
  iframe.style.transformOrigin = 'top left'
  wrapper.style.overflow = 'hidden'
}

window.addEventListener('resize', () => {
  if (state.slides.length > 0) scalePreview()
})

// ── Navigation ────────────────────────────────────────────────────────────

function navigate(delta) {
  const newIdx = state.currentIndex + delta
  if (newIdx < 0 || newIdx >= state.slides.length) return
  goToSlide(newIdx)
}

function goToSlide(index) {
  if (index < 0 || index >= state.slides.length) return

  // Flush any unsaved editor content to the current slide before switching
  if (state.editorOpen && index !== state.currentIndex) {
    const html = getEditorContent()
    if (html && html !== state.slides[state.currentIndex].content) {
      state.slides[state.currentIndex] = { ...state.slides[state.currentIndex], content: html }
      setDirty(true)
      refreshThumb(state.currentIndex)
    }
  }

  state.currentIndex = index
  updateThumbnailActive()
  renderPreview()
}

function updateNavButtons() {
  const prevBtn = document.getElementById('prev-btn')
  const nextBtn = document.getElementById('next-btn')
  prevBtn.disabled = state.currentIndex <= 0
  nextBtn.disabled = state.currentIndex >= state.slides.length - 1
}

function updatePageInfo() {
  document.getElementById('page-info').textContent =
    state.slides.length > 0
      ? `${state.currentIndex + 1} / ${state.slides.length}`
      : '0 / 0'
}

// ── Notes Panel ───────────────────────────────────────────────────────────

function updateNotesPanel() {
  const textarea = document.getElementById('notes-content')
  if (!textarea) return
  const slide = state.slides[state.currentIndex]
  textarea.value = slide ? (slide.notes || '') : ''
}

function setupNotesPanel() {
  const toggle = document.getElementById('notes-toggle')
  const panel = document.getElementById('notes-panel')
  const textarea = document.getElementById('notes-content')
  if (!toggle || !panel || !textarea) return

  toggle.addEventListener('click', () => {
    panel.classList.toggle('collapsed')
  })

  textarea.addEventListener('input', () => {
    if (state.slides.length === 0) return
    state.slides[state.currentIndex] = {
      ...state.slides[state.currentIndex],
      notes: textarea.value
    }
    setDirty(true)
  })
}

// ── Edit Mode ─────────────────────────────────────────────────────────────

function toggleEditMode() {
  if (state.editMode === 'source') {
    // Close editor and return to visual
    closeEditor()
    const iframe = document.getElementById('main-preview')
    // Only activate visual edit if not in preview mode
    if (!state.previewMode) {
      activateVisualEdit(iframe)
    }
  } else {
    // Enter source mode
    state.editMode = 'source'
    document.getElementById('edit-mode-btn').classList.add('active')
    openEditorDrawer()
    const iframe = document.getElementById('main-preview')
    if (state.visualEditActive) {
      disableVisualEdit(iframe)
      state.visualEditActive = false
    }
  }
}

function openEditorDrawer() {
  const drawer = document.getElementById('editor-drawer')
  drawer.classList.add('open')
  drawer.style.height = state.editorHeight + 'px'
  state.editorOpen = true
  updateEditorContent()
  // Resize preview
  setTimeout(scalePreview, 250)
}

function closeEditor() {
  const drawer = document.getElementById('editor-drawer')
  drawer.classList.remove('open')
  drawer.style.height = '0'
  state.editorOpen = false
  state.editMode = 'visual'
  document.getElementById('edit-mode-btn').classList.remove('active')
  setTimeout(scalePreview, 250)
}

function updateEditorContent() {
  if (!state.editorOpen || state.slides.length === 0) return
  const slide = state.slides[state.currentIndex]
  setEditorContent(slide.content)
  const label = document.getElementById('editor-slide-label')
  if (label) label.textContent = `幻灯片 ${state.currentIndex + 1}：${slide.title || ''}`
}

function activateVisualEdit(iframe) {
  if (!iframe || !iframe.contentDocument) return
  state.visualEditActive = true
  enableVisualEdit(iframe, (fullHTML) => {
    applySlideContent(fullHTML, true)
  })
}

// ── Slide Content Changes ─────────────────────────────────────────────────

function applySlideContent(html, pushUndo = true) {
  if (state.slides.length === 0) return

  const idx = state.currentIndex
  if (pushUndo) pushUndoState(idx)

  state.slides[idx] = { ...state.slides[idx], content: html }
  setDirty(true)

  refreshThumb(idx)
  // Don't reload the iframe if it's from visual edit (already reflected in DOM)
  if (!state.visualEditActive) {
    renderPreview()
  }

  if (state.editorOpen) {
    setEditorContent(html)
  }
}

function replaceCurrentSlide(html) {
  if (state.slides.length === 0) {
    // No existing slides — create new slide
    loadHTMLContent(html, null)
    return
  }
  pushUndoState(state.currentIndex)
  state.slides[state.currentIndex] = {
    ...state.slides[state.currentIndex],
    content: html
  }
  setDirty(true)
  refreshThumb(state.currentIndex)
  renderPreview()
}

// ── Undo / Redo ───────────────────────────────────────────────────────────

function pushUndoState(index) {
  if (!state.undoStacks[index]) state.undoStacks[index] = []
  state.undoStacks[index].push(state.slides[index].content)
  if (state.undoStacks[index].length > 50) state.undoStacks[index].shift()
  // Clear redo on new change
  state.redoStacks[index] = []
  updateUndoRedoButtons()
}

function undo() {
  const idx = state.currentIndex
  const stack = state.undoStacks[idx]
  if (!stack || stack.length === 0) return

  if (!state.redoStacks[idx]) state.redoStacks[idx] = []
  state.redoStacks[idx].push(state.slides[idx].content)

  const html = stack.pop()
  state.slides[idx] = { ...state.slides[idx], content: html }
  setDirty(true)
  refreshThumb(idx)
  renderPreview()
  updateUndoRedoButtons()
}

function redo() {
  const idx = state.currentIndex
  const stack = state.redoStacks[idx]
  if (!stack || stack.length === 0) return

  if (!state.undoStacks[idx]) state.undoStacks[idx] = []
  state.undoStacks[idx].push(state.slides[idx].content)

  const html = stack.pop()
  state.slides[idx] = { ...state.slides[idx], content: html }
  setDirty(true)
  refreshThumb(idx)
  renderPreview()
  updateUndoRedoButtons()
}

function updateUndoRedoButtons() {
  const idx = state.currentIndex
  const canUndo = !!(state.undoStacks[idx]?.length > 0)
  const canRedo = !!(state.redoStacks[idx]?.length > 0)
  const undoBtn = document.getElementById('undo-btn')
  const redoBtn = document.getElementById('redo-btn')
  undoBtn.disabled = !canUndo
  undoBtn.classList.toggle('active', canUndo)
  redoBtn.disabled = !canRedo
  redoBtn.classList.toggle('active', canRedo)
}

// ── UI Toggles ────────────────────────────────────────────────────────────

function toggleSidebar() {
  state.sidebarOpen = !state.sidebarOpen
  document.getElementById('sidebar').classList.toggle('collapsed', !state.sidebarOpen)
  document.getElementById('sidebar-toggle-btn').classList.toggle('active', !state.sidebarOpen)
  setTimeout(scalePreview, 250)
}

function toggleAIPanel() {
  state.aiPanelOpen = !state.aiPanelOpen
  document.getElementById('ai-panel').classList.toggle('collapsed', !state.aiPanelOpen)
  document.getElementById('ai-toggle-btn').classList.toggle('active', state.aiPanelOpen)
  setTimeout(scalePreview, 250)
}

// ── Presentation ──────────────────────────────────────────────────────────

async function openPresentation() {
  if (state.slides.length === 0) return

  const startIdx = state.currentIndex
  const total = state.slides.length

  // Build each slide as an iframe[srcdoc] so CSS is fully isolated
  const slideFrames = state.slides.map((slide, i) => {
    const content = slide.content
    const encoded = content.replace(/&/g, '&amp;').replace(/"/g, '&quot;')
    const cls = i === startIdx ? ' class="active"' : ''
    return `<iframe${cls} sandbox="allow-scripts allow-same-origin" scrolling="no" srcdoc="${encoded}"></iframe>`
  }).join('\n')

  const presHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<style>
*{margin:0;padding:0;box-sizing:border-box}
html,body{width:100%;height:100%;background:#000;overflow:hidden;
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif}

/* ── Stage: centered 1280×720 that scales to fill screen ── */
#stage{
  position:fixed;top:50%;left:50%;
  width:1280px;height:720px;
  transform:translate(-50%,-50%) scale(var(--s,1));
  transform-origin:center center;
}
#stage iframe{
  position:absolute;inset:0;
  width:1280px;height:720px;
  border:none;display:none;background:#fff;
}
#stage iframe.active{display:block}

/* ── Always-visible exit pill (top-right) ── */
#exit-pill{
  position:fixed;top:16px;right:16px;z-index:9999;
  background:rgba(0,0,0,0.55);border:1px solid rgba(255,255,255,0.2);
  color:rgba(255,255,255,0.6);padding:5px 14px;border-radius:20px;
  font-size:12px;cursor:pointer;transition:all .2s;
}
#exit-pill:hover{background:rgba(200,40,40,.9);color:#fff;border-color:transparent}

/* ── Progress bar ── */
#progress{
  position:fixed;bottom:0;left:0;height:3px;
  background:rgba(230,126,34,.85);transition:width .3s;z-index:9999;
}

/* ── HUD (shows on mouse move) ── */
#hud{
  position:fixed;bottom:28px;left:50%;transform:translateX(-50%);z-index:9999;
  background:rgba(0,0,0,.72);backdrop-filter:blur(10px);
  border:1px solid rgba(255,255,255,.15);border-radius:32px;
  padding:10px 22px;display:flex;align-items:center;gap:14px;
  color:#fff;font-size:14px;
  opacity:0;transition:opacity .35s;pointer-events:none;
}
body.hud-on #hud{opacity:1;pointer-events:auto}
.hbtn{
  background:none;border:1px solid rgba(255,255,255,.3);color:#fff;
  padding:5px 16px;border-radius:16px;cursor:pointer;font-size:13px;transition:background .15s;
}
.hbtn:hover{background:rgba(255,255,255,.18)}
.hbtn.exit{background:rgba(210,40,40,.8);border-color:transparent;padding:5px 12px}
.hbtn.exit:hover{background:rgba(210,40,40,1)}
#counter{min-width:52px;text-align:center;font-size:15px;font-weight:600}
</style>
</head>
<body>

<div id="stage">${slideFrames}</div>

<!-- always-visible exit -->
<button id="exit-pill" onclick="window.close()">ESC 退出演示</button>

<!-- progress bar -->
<div id="progress" style="width:${(1/total*100).toFixed(1)}%"></div>

<!-- HUD -->
<div id="hud">
  <button class="hbtn" id="btn-prev" onclick="prev()">◀ 上一页</button>
  <span id="counter">${startIdx+1} / ${total}</span>
  <button class="hbtn" id="btn-next" onclick="next()">下一页 ▶</button>
  <button class="hbtn exit" onclick="window.close()">✕ 退出</button>
</div>

<script>
var cur=${startIdx}, tot=${total};
var frames=document.querySelectorAll('#stage iframe');
var hudTimer=null;

function show(n){
  n=Math.max(0,Math.min(n,frames.length-1));
  cur=n;
  frames.forEach(function(f,i){f.classList.toggle('active',i===n)});
  document.getElementById('counter').textContent=(n+1)+' / '+frames.length;
  document.getElementById('progress').style.width=((n+1)/frames.length*100).toFixed(1)+'%';
  document.getElementById('btn-prev').disabled=(n===0);
  document.getElementById('btn-next').disabled=(n===frames.length-1);
}

function next(){if(cur<frames.length-1)show(cur+1)}
function prev(){if(cur>0)show(cur-1)}

function showHUD(){
  document.body.classList.add('hud-on');
  clearTimeout(hudTimer);
  hudTimer=setTimeout(function(){document.body.classList.remove('hud-on')},3000);
}

document.addEventListener('mousemove',showHUD);
document.addEventListener('keydown',function(e){
  showHUD();
  if(e.key==='ArrowRight'||e.key===' '||e.key==='ArrowDown'||e.key==='PageDown'){e.preventDefault();next()}
  else if(e.key==='ArrowLeft'||e.key==='ArrowUp'||e.key==='PageUp'){e.preventDefault();prev()}
  else if(e.key==='Escape')window.close();
});
document.addEventListener('click',function(e){
  if(!e.target.closest('#hud')&&!e.target.closest('#exit-pill'))next();
});

function scale(){
  var sx=window.innerWidth/1280,sy=window.innerHeight/720;
  document.getElementById('stage').style.setProperty('--s',Math.min(sx,sy));
}
window.addEventListener('resize',scale);
scale();
show(cur);
showHUD();
<\/script>
</body></html>`

  await window.electronAPI.openPresentation(presHtml, startIdx)
}

// ── Dirty / Title ─────────────────────────────────────────────────────────

function setDirty(dirty) {
  state.isDirty = dirty
  window.electronAPI.setDocumentEdited(dirty)
  window.electronAPI.setDirtyFlag(dirty)
  updateTitle()
  updateUndoRedoButtons()
}

function updateTitle() {
  const base = state.filePath
    ? state.filePath.split('/').pop().split('\\').pop()
    : state.slides.length > 0 ? '未命名.html' : 'Slide X'
  const prefix = state.isDirty ? '● ' : ''
  window.electronAPI.setTitle(`${prefix}${base} — Slide X`)
}

function showWelcome(show) {
  document.getElementById('welcome').classList.toggle('hidden', !show)
}

// ── Bootstrap ────────────────────────────────────────────────────────────

initApp().catch(err => {
  console.error('App initialization failed:', err)
  // Show error to user
  const welcome = document.getElementById('welcome')
  if (welcome) {
    welcome.innerHTML = `
      <div id="welcome-icon">⚠️</div>
      <h2>初始化失败</h2>
      <p>${err.message || '未知错误'}<br><br>请刷新页面重试</p>
    `
  }
})
