/**
 * ai-panel.js — AI generation panel with Memory, Style, and Intent detection
 */

import { STYLE_TEMPLATES, getTemplateById, buildStylePrompt } from './style-templates.js'
import { detectIntent } from './intent-detector.js'

let currentAbortController = null
let onNewPPT = null
let onModifySlide = null
let getState = null

// ── System Prompts ──────────────────────────────────────────────────────────

const SYSTEM_PROMPT_FULL = `你是一个专业的 PPT 设计师。请生成一个完整的 HTML PPT 文件。

严格遵守以下结构规范，每页用 <section data-slide="N" data-title="页面标题"> 包裹：

<!DOCTYPE html>
<html>
<head>
<style>
  body { margin: 0; padding: 0; }
  section[data-slide] {
    width: 1280px;
    height: 720px;
    position: relative;
    overflow: hidden;
    display: none;
  }
</style>
</head>
<body>
  <section data-slide="1" data-title="封面">...</section>
  <section data-slide="2" data-title="...">...</section>
</body>
</html>

## 设计规范（严格遵守）

### 配色原则
- 选择与主题内容高度匹配的配色，绝对不要默认使用通用蓝色
- 一种主色应占 60-70% 视觉权重，1-2 种辅色，一种强调色
- 封面和结语页使用深色背景，内容页可使用浅色背景（"三明治"结构）或全程深色
- 主色、辅色、强调色形成强烈对比

### 布局多样性
- 不同页面使用不同布局，不要重复相同的排版
- 可用布局：双栏（文字左/图示右）、图标+文字行、2x2或2x3网格卡片、半出血图片+内容叠加
- 数据展示：大号数字统计（60-72pt 数字+小标签）、对比栏、时间线/流程图

### 视觉元素
- 每页必须有视觉元素（色块、几何图形、渐变、图标符号等）
- 使用 CSS 渐变、几何形状、emoji 图标作为视觉辅助
- 标题字号 44-52px，正文 16-20px，说明文字 12-14px

### 绝对禁止
- ❌ 标题下方画横线/装饰线（这是 AI 生成幻灯片的明显特征）
- ❌ 纯文字页面（每页必须有视觉元素）
- ❌ 正文居中对齐（列表和段落必须左对齐）
- ❌ 所有页面使用相同布局
- ❌ 引用任何外部资源（无外链字体、图片用 CSS 渐变代替）
- ❌ 低对比度文字（浅色背景上不用浅灰文字）

### 技术要求
- 每页固定尺寸 1280×720px，position: relative，overflow: hidden
- 使用内联 CSS，style 写在 <head> 或元素上
- 只输出完整 HTML 代码，不要任何解释文字`

const SYSTEM_PROMPT_MODIFY = `你是一个专业的前端开发者，擅长修改 HTML PPT 幻灯片。
用户会给你一段幻灯片的 HTML 代码，以及修改指令。
请按指令修改 HTML，保持外层结构不变，宽度保持 1280px，高度保持 720px。
只输出修改后的完整 HTML 代码，不要任何解释。`

// ── Init ────────────────────────────────────────────────────────────────────

export async function initAIPanel({ onGenerate, onModify, getAppState }) {
  onNewPPT = onGenerate
  onModifySlide = onModify
  getState = getAppState

  // Inject panel HTML
  renderAIPanel()

  // Tab switching
  document.getElementById('ai-tab-generate').addEventListener('click', () => switchTab('generate'))
  document.getElementById('ai-tab-modify').addEventListener('click', () => switchTab('modify'))
  document.getElementById('ai-tab-memory').addEventListener('click', () => switchTab('memory'))
  document.getElementById('ai-tab-style').addEventListener('click', () => switchTab('style'))

  // Generate
  document.getElementById('ai-generate-btn').addEventListener('click', handleGenerate)

  // Modify
  document.getElementById('ai-modify-btn').addEventListener('click', handleModify)

  // Stop
  document.getElementById('ai-stop-btn').addEventListener('click', stopGeneration)

  // Settings
  document.getElementById('settings-btn').addEventListener('click', openSettings)
  document.getElementById('settings-close-btn').addEventListener('click', closeSettings)
  document.getElementById('settings-save-btn').addEventListener('click', saveSettings)
  document.getElementById('toggle-api-key').addEventListener('click', toggleApiKeyVisibility)

  // Memory
  document.getElementById('memory-upload-btn').addEventListener('click', handleMemoryUpload)
  const dropZone = document.getElementById('memory-drop-zone')
  dropZone.addEventListener('dragover', onMemoryDragOver)
  dropZone.addEventListener('dragleave', (e) => { e.currentTarget.classList.remove('drag-over') })
  dropZone.addEventListener('drop', onMemoryDrop)

  // Event delegation for memory list actions (avoids global window exposure)
  document.getElementById('memory-list').addEventListener('click', (e) => {
    const btn = e.target.closest('[data-action]')
    if (!btn) return
    const item = btn.closest('.memory-item')
    const fileId = item?.dataset.id
    if (!fileId) return
    if (btn.dataset.action === 'toggle') toggleMemorySelection(fileId)
    else if (btn.dataset.action === 'delete') deleteMemoryFile(fileId)
  })

  // Style
  initStylePanel()

  // Intent detection on topic input
  const topicInput = document.getElementById('ai-topic')
  topicInput.addEventListener('input', debounce(handleTopicIntentHint, 400))

  // Intent detection on modify input
  const modifyInput = document.getElementById('ai-modify-instruction')
  modifyInput.addEventListener('input', debounce(handleModifyIntentHint, 400))

  // Load config once and initialize settings + style + memory
  const config = await window.electronAPI.getConfig()
  applyConfigToForms(config)
  loadMemoryList()
}

// ── Tab Switching ───────────────────────────────────────────────────────────

function switchTab(tab) {
  const tabs = ['generate', 'modify', 'memory', 'style']
  tabs.forEach(t => {
    const btn = document.getElementById(`ai-tab-${t}`)
    const form = document.getElementById(`ai-${t}-form`)
    if (btn) btn.classList.toggle('active', t === tab)
    if (form) form.style.display = t === tab ? 'flex' : 'none'
  })
}

// ── Intent Detection ────────────────────────────────────────────────────────

function handleTopicIntentHint() {
  const input = document.getElementById('ai-topic').value
  if (!input.trim()) {
    clearIntentHint('generate')
    return
  }
  const state = getState()
  const result = detectIntent(input, state?.slides?.length > 0)

  if (result.intent === 'modify' && result.confidence > 0.6) {
    showIntentHint('generate', '💡 ' + result.suggestion + '，或直接在"修改当前页"中输入', 'info')
    // Auto-populate modify textarea
    document.getElementById('ai-modify-instruction').value = input
  } else {
    clearIntentHint('generate')
  }
}

function handleModifyIntentHint() {
  const input = document.getElementById('ai-modify-instruction').value
  if (!input.trim()) {
    clearIntentHint('modify')
    return
  }
  const state = getState()
  const result = detectIntent(input, state?.slides?.length > 0)

  if (result.intent === 'generate' && result.confidence > 0.7) {
    showIntentHint('modify', '💡 ' + result.suggestion + '，或在"全新生成"中输入主题', 'info')
  } else {
    clearIntentHint('modify')
  }
}

function showIntentHint(tab, text, type = 'info') {
  const el = document.getElementById(`intent-hint-${tab}`)
  if (!el) return
  el.textContent = text
  el.className = `intent-hint intent-hint-${type} visible`
}

function clearIntentHint(tab) {
  const el = document.getElementById(`intent-hint-${tab}`)
  if (!el) return
  el.className = 'intent-hint'
  el.textContent = ''
}

// ── Generate ────────────────────────────────────────────────────────────────

async function handleGenerate() {
  setGenerating(true)

  const config = await window.electronAPI.getConfig()
  if (!config.apiKey) {
    setGenerating(false)
    alert('请先在设置中配置 API Key')
    openSettings()
    return
  }

  const topic = document.getElementById('ai-topic').value.trim()
  if (!topic) {
    setGenerating(false)
    document.getElementById('ai-topic').focus()
    return
  }

  const pages = Math.max(3, Math.min(30, parseInt(document.getElementById('ai-pages').value) || 8))
  const lang = document.getElementById('ai-lang').value

  const state = getState()
  if (state.isDirty && state.slides.length > 0) {
    const result = await window.electronAPI.showMessageBox({
      type: 'question',
      buttons: ['保存', '不保存，继续', '取消'],
      defaultId: 1,
      cancelId: 2,
      message: '当前文件有未保存的修改',
      detail: '是否在生成前保存？'
    })
    if (result.response === 0) {
      document.dispatchEvent(new CustomEvent('app:save'))
      await new Promise(r => setTimeout(r, 300))
    } else if (result.response === 2) {
      setGenerating(false)
      return
    }
  }

  // Build prompt with memory and style
  const memoryContent = await getSelectedMemoryContent()
  const styleId = document.getElementById('style-template-select')?.value || ''
  const styleParams = getStyleParams()
  const userPrompt = buildGeneratePrompt(topic, pages, lang, memoryContent, styleId, styleParams)
  const systemPrompt = buildSystemPromptWithStyle(styleId)

  const maxTokens = 128000

  showProgress('正在生成 PPT...')

  try {
    const html = await streamCompletion(config, systemPrompt, userPrompt, (chunk, total) => {
      showProgress(`正在生成... ${total} 字符`, chunk.slice(-200))
    }, maxTokens)

    const clean = extractHTML(html)
    if (onNewPPT) onNewPPT(clean)
    showProgress('✓ 生成完成！', '')
    setTimeout(() => hideProgress(), 2000)
  } catch (err) {
    if (err.name !== 'AbortError') {
      showProgress('✗ 生成失败：' + err.message)
      setTimeout(() => hideProgress(), 5000)
    } else {
      showProgress('已停止')
      setTimeout(() => hideProgress(), 1500)
    }
  } finally {
    setGenerating(false)
  }
}

// ── Modify ──────────────────────────────────────────────────────────────────

async function handleModify() {
  setGenerating(true, true)

  const config = await window.electronAPI.getConfig()
  if (!config.apiKey) {
    setGenerating(false, true)
    alert('请先在设置中配置 API Key')
    openSettings()
    return
  }

  const instruction = document.getElementById('ai-modify-instruction').value.trim()
  if (!instruction) {
    setGenerating(false, true)
    document.getElementById('ai-modify-instruction').focus()
    return
  }

  const state = getState()
  if (!state.slides || state.slides.length === 0) {
    setGenerating(false, true)
    alert('请先打开或生成一个 PPT')
    return
  }

  const currentSlide = state.slides[state.currentIndex]

  // Build context: include adjacent slides and global style info
  const contextInfo = buildModifyContext(state)
  const memoryContent = await getSelectedMemoryContent()

  const userPrompt = buildModifyPrompt(currentSlide.content, instruction, contextInfo, memoryContent)

  showProgress('正在修改幻灯片...')

  try {
    const html = await streamCompletion(config, SYSTEM_PROMPT_MODIFY, userPrompt, (chunk, total) => {
      showProgress(`正在修改... ${total} 字符`)
    })

    const clean = extractHTML(html)
    if (onModifySlide) onModifySlide(clean)
    showProgress('✓ 修改完成！')
    setTimeout(() => hideProgress(), 2000)
  } catch (err) {
    if (err.name !== 'AbortError') {
      showProgress('✗ 修改失败：' + err.message)
      setTimeout(() => hideProgress(), 5000)
    } else {
      showProgress('已停止')
      setTimeout(() => hideProgress(), 1500)
    }
  } finally {
    setGenerating(false, true)
  }
}

function buildModifyContext(state) {
  const slides = state.slides
  const idx = state.currentIndex
  const lines = []

  // Extract global <style> from first slide's parent context
  if (slides.length > 0 && slides[0].content) {
    const styleMatch = slides[0].content.match(/<style[^>]*>([\s\S]*?)<\/style>/i)
    if (styleMatch) {
      lines.push(`全局样式参考（保持一致）：\n<style>${styleMatch[1].slice(0, 2000)}</style>`)
    }
  }

  // Adjacent slides for context
  if (idx > 0) {
    const prev = slides[idx - 1]
    lines.push(`上一页（第${idx}页）内容摘要：\n${extractSlideText(prev.content)}`)
  }
  if (idx < slides.length - 1) {
    const next = slides[idx + 1]
    lines.push(`下一页（第${idx + 2}页）内容摘要：\n${extractSlideText(next.content)}`)
  }

  lines.push(`当前页为第 ${idx + 1} 页，共 ${slides.length} 页`)

  return lines.join('\n\n')
}

function extractSlideText(html) {
  if (!html) return ''
  // Strip tags and get visible text, limit to 300 chars
  return html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 300)
}

// ── Prompt Builders ─────────────────────────────────────────────────────────

function buildGeneratePrompt(topic, pages, lang, memoryContent, styleId, styleParams) {
  let prompt = `请生成一个关于「${topic}」的完整 PPT，严格要求 ${pages} 页。`
  prompt += `\n必须生成从 data-slide="1" 到 data-slide="${pages}" 的全部 ${pages} 个 <section>，禁止在未完成所有页面前结束输出。`

  if (pages >= 20) {
    prompt += `\n页数较多，请保持每页 HTML 精简（避免过多嵌套和内联 style 冗余），确保全部 ${pages} 页都能输出完整。`
  }

  if (lang === 'en') prompt += '\n请用英文生成所有文字内容。'
  else prompt += '\n请用中文生成所有文字内容。'

  // Append style instructions
  if (styleId) {
    const styleHint = buildStylePrompt(styleId, styleParams)
    if (styleHint) prompt += styleHint
  }

  // Append memory content as background knowledge
  if (memoryContent) {
    prompt += `\n\n以下是相关背景知识，请据此生成更准确、专业的内容：\n\n${memoryContent}`
  }

  return prompt
}

function buildSystemPromptWithStyle(styleId) {
  const template = getTemplateById(styleId)
  if (!template) return SYSTEM_PROMPT_FULL

  return SYSTEM_PROMPT_FULL + `\n\n风格配色参考（请严格遵守）：
主色 ${template.colors.primary}，辅色 ${template.colors.secondary}，强调色 ${template.colors.accent}
背景色 ${template.colors.background}，主文字色 ${template.colors.text}，次要文字色 ${template.colors.textMuted}
标题字重：${template.fonts.title}，正文字重：${template.fonts.body}`
}

function buildModifyPrompt(slideHtml, instruction, contextInfo, memoryContent) {
  let prompt = `以下是当前幻灯片的 HTML：\n\n\`\`\`html\n${slideHtml}\n\`\`\``

  if (contextInfo) {
    prompt += `\n\n上下文参考信息（仅用于保持风格一致性，不要修改这些页面）：\n${contextInfo}`
  }

  if (memoryContent) {
    prompt += `\n\n背景知识参考：\n${memoryContent}`
  }

  prompt += `\n\n修改指令：${instruction}`
  return prompt
}

// ── Memory ──────────────────────────────────────────────────────────────────

let memoryFiles = []

async function loadMemoryList() {
  try {
    memoryFiles = await window.electronAPI.getMemoryList()
    renderMemoryList()
  } catch (e) {
    console.error('Failed to load memory list:', e)
  }
}

async function handleMemoryUpload() {
  const result = await window.electronAPI.showMemoryFileDialog()
  if (result.canceled || !result.filePaths.length) return
  await uploadFiles(result.filePaths)
}

function onMemoryDragOver(e) {
  e.preventDefault()
  e.currentTarget.classList.add('drag-over')
}

async function onMemoryDrop(e) {
  e.preventDefault()
  e.currentTarget.classList.remove('drag-over')
  const files = Array.from(e.dataTransfer.files).map(f => f.path).filter(Boolean)
  if (files.length) await uploadFiles(files)
}

async function uploadFiles(filePaths) {
  const progressEl = document.getElementById('memory-upload-progress')
  progressEl.style.display = 'block'
  progressEl.textContent = '解析中...'

  let successCount = 0
  const errors = []

  for (const filePath of filePaths) {
    try {
      progressEl.textContent = `解析 ${filePath.split('/').pop()}...`
      const fileEntry = await window.electronAPI.parseMemoryFile(filePath)
      await window.electronAPI.saveMemoryFile(fileEntry)
      memoryFiles.push(fileEntry)
      successCount++
    } catch (e) {
      errors.push(`${filePath.split('/').pop()}: ${e?.message || String(e)}`)
    }
  }

  renderMemoryList()

  if (errors.length > 0) {
    progressEl.textContent = `完成 ${successCount} 个，失败 ${errors.length} 个: ${errors.join('; ')}`
  } else {
    progressEl.textContent = `✓ 成功添加 ${successCount} 个文件`
    setTimeout(() => { progressEl.style.display = 'none' }, 2500)
  }
}

async function deleteMemoryFile(fileId) {
  await window.electronAPI.deleteMemoryFile(fileId)
  memoryFiles = memoryFiles.filter(f => f.id !== fileId)
  renderMemoryList()
}

async function toggleMemorySelection(fileId) {
  const file = memoryFiles.find(f => f.id === fileId)
  if (!file) return
  file.selected = !file.selected
  await window.electronAPI.saveMemoryFile(file)
  renderMemoryList()
}

function renderMemoryList() {
  const listEl = document.getElementById('memory-list')
  if (!listEl) return

  if (memoryFiles.length === 0) {
    listEl.innerHTML = '<div class="memory-empty">暂无附件<br>上传文档作为 AI 的背景知识</div>'
    updateMemoryBadge()
    return
  }

  listEl.innerHTML = memoryFiles.map(file => `
    <div class="memory-item ${file.selected ? 'selected' : ''}" data-id="${file.id}">
      <div class="memory-item-icon">${getFileIcon(file.type)}</div>
      <div class="memory-item-info">
        <div class="memory-item-name">${escapeHtml(file.name)}</div>
        <div class="memory-item-meta">${file.type.toUpperCase()} · ${formatSize(file.size)} · ${file.content.length} 字符</div>
      </div>
      <div class="memory-item-actions">
        <button class="memory-toggle-btn ${file.selected ? 'active' : ''}" data-action="toggle" title="${file.selected ? '取消使用' : '启用此文件'}">${file.selected ? '✓ 启用' : '启用'}</button>
        <button class="memory-delete-btn" data-action="delete" title="删除">✕</button>
      </div>
    </div>
  `).join('')

  updateMemoryBadge()
}

function updateMemoryBadge() {
  const count = memoryFiles.filter(f => f.selected).length
  const badge = document.getElementById('memory-badge')
  if (badge) {
    badge.textContent = count > 0 ? count : ''
    badge.style.display = count > 0 ? 'inline-block' : 'none'
  }
}

async function getSelectedMemoryContent() {
  const selected = memoryFiles.filter(f => f.selected)
  if (selected.length === 0) return ''

  const MAX_TOTAL = 50000
  let combined = ''

  for (const file of selected) {
    const header = `\n--- ${file.name} ---\n`
    const available = MAX_TOTAL - combined.length - header.length
    if (available <= 100) break
    combined += header + file.content.slice(0, available)
  }

  return combined.trim()
}

// ── Style Panel ─────────────────────────────────────────────────────────────

function initStylePanel() {
  const selectEl = document.getElementById('style-template-select')
  if (!selectEl) return

  // Render template options
  selectEl.innerHTML = '<option value="">无（自定义描述）</option>' +
    STYLE_TEMPLATES.map(t =>
      `<option value="${t.id}">${t.emoji} ${t.name} — ${t.description}</option>`
    ).join('')

  selectEl.addEventListener('change', updateStylePreview)
  document.getElementById('style-color-temp').addEventListener('input', updateStylePreview)
  document.getElementById('style-contrast').addEventListener('input', updateStylePreview)
  document.getElementById('style-density').addEventListener('input', updateStylePreview)
  document.getElementById('style-save-btn').addEventListener('click', saveStyleConfig)
}

function applyConfigToForms(config) {
  // Settings form
  document.getElementById('settings-base-url').value = config.baseUrl || 'https://api.openai.com/v1'
  document.getElementById('settings-api-key').value = config.apiKey || ''
  document.getElementById('settings-model').value = config.model || 'gpt-4o'

  // Style form
  const selectEl = document.getElementById('style-template-select')
  if (selectEl && config.styleConfig) {
    if (config.styleConfig.templateId) selectEl.value = config.styleConfig.templateId
    if (config.styleConfig.colorTemp !== undefined) {
      document.getElementById('style-color-temp').value = config.styleConfig.colorTemp
    }
    if (config.styleConfig.contrast !== undefined) {
      document.getElementById('style-contrast').value = config.styleConfig.contrast
    }
    if (config.styleConfig.density !== undefined) {
      document.getElementById('style-density').value = config.styleConfig.density
    }
  }
  updateStylePreview()
}

function getStyleParams() {
  return {
    colorTemp: parseInt(document.getElementById('style-color-temp')?.value || '50'),
    contrast: parseInt(document.getElementById('style-contrast')?.value || '50'),
    density: parseInt(document.getElementById('style-density')?.value || '50')
  }
}

function updateStylePreview() {
  const templateId = document.getElementById('style-template-select')?.value
  const previewEl = document.getElementById('style-preview')
  if (!previewEl) return

  const template = getTemplateById(templateId)
  if (!template) {
    previewEl.innerHTML = '<div class="style-preview-empty">选择模板预览效果</div>'
    return
  }

  const params = getStyleParams()
  const bg = template.colors.background
  const primary = template.colors.primary
  const text = template.colors.text

  previewEl.innerHTML = `
    <div class="style-preview-card" style="background:${bg};color:${text};border-color:${primary}30;">
      <div class="style-preview-title" style="color:${primary};font-weight:${template.fonts.title};">
        ${template.emoji} ${template.name}
      </div>
      <div class="style-preview-subtitle" style="color:${template.colors.secondary};">副标题示例文本</div>
      <div class="style-preview-body" style="color:${template.colors.textMuted};">
        正文内容示例，展示整体配色和排版效果
      </div>
      <div class="style-preview-accent" style="background:${template.colors.accent};"></div>
      <div class="style-preview-hint">${template.promptHint.slice(0, 50)}...</div>
    </div>
    <div class="style-params-display">
      色温 ${params.colorTemp}% · 对比度 ${params.contrast}% · 密度 ${params.density}%
    </div>
  `
}

async function saveStyleConfig() {
  const templateId = document.getElementById('style-template-select')?.value || ''
  const params = getStyleParams()
  await window.electronAPI.setConfig({
    styleConfig: { templateId, ...params }
  })
  const saveBtn = document.getElementById('style-save-btn')
  const orig = saveBtn.textContent
  saveBtn.textContent = '✓ 已保存'
  setTimeout(() => { saveBtn.textContent = orig }, 1500)
}

// ── Streaming ───────────────────────────────────────────────────────────────

function stopGeneration() {
  if (currentAbortController) {
    currentAbortController.abort()
    currentAbortController = null
  }
}

async function streamCompletion(config, systemPrompt, userPrompt, onChunk, maxTokens = 16384) {
  currentAbortController = new AbortController()

  const baseUrl = (config.baseUrl || 'https://api.openai.com/v1').replace(/\/$/, '')
  const model = config.model || 'gpt-4o'
  const apiKey = config.apiKey

  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      stream: true,
      max_tokens: maxTokens
    }),
    signal: currentAbortController.signal
  })

  if (!response.ok) {
    const err = await response.text()
    throw new Error(`API Error ${response.status}: ${err}`)
  }

  const reader = response.body.getReader()
  const decoder = new TextDecoder()
  let buffer = ''
  let fullContent = ''

  while (true) {
    const { done, value } = await reader.read()
    if (done) break

    buffer += decoder.decode(value, { stream: true })
    const lines = buffer.split('\n')
    buffer = lines.pop()

    for (const line of lines) {
      const trimmed = line.trim()
      if (!trimmed || !trimmed.startsWith('data: ')) continue
      const data = trimmed.slice(6)
      if (data === '[DONE]') continue
      try {
        const parsed = JSON.parse(data)
        const delta = parsed.choices?.[0]?.delta?.content || ''
        if (delta) {
          fullContent += delta
          onChunk(delta, fullContent.length)
        }
      } catch (e) {
        // Skip malformed chunks
      }
    }
  }

  return fullContent
}

// ── Settings ─────────────────────────────────────────────────────────────────

async function loadSettingsToForm() {
  const config = await window.electronAPI.getConfig()
  document.getElementById('settings-base-url').value = config.baseUrl || 'https://api.openai.com/v1'
  document.getElementById('settings-api-key').value = config.apiKey || ''
  document.getElementById('settings-model').value = config.model || 'gpt-4o'
}

async function saveSettings() {
  const baseUrl = document.getElementById('settings-base-url').value.trim()
  const apiKey = document.getElementById('settings-api-key').value.trim()
  const model = document.getElementById('settings-model').value.trim()
  await window.electronAPI.setConfig({ baseUrl, apiKey, model })
  closeSettings()
}

function openSettings() {
  loadSettingsToForm()
  document.getElementById('settings-modal').classList.remove('hidden')
}

function closeSettings() {
  document.getElementById('settings-modal').classList.add('hidden')
}

function toggleApiKeyVisibility() {
  const input = document.getElementById('settings-api-key')
  const btn = document.getElementById('toggle-api-key')
  if (input.type === 'password') {
    input.type = 'text'
    btn.textContent = '隐藏'
  } else {
    input.type = 'password'
    btn.textContent = '显示'
  }
}

// ── UI Helpers ───────────────────────────────────────────────────────────────

function setGenerating(active, isModify = false) {
  const genBtn = document.getElementById('ai-generate-btn')
  const modBtn = document.getElementById('ai-modify-btn')
  const stopBtn = document.getElementById('ai-stop-btn')

  if (isModify) {
    modBtn.disabled = active
  } else {
    genBtn.disabled = active
  }
  stopBtn.classList.toggle('visible', active)
}

function showProgress(text, detail) {
  const el = document.getElementById('ai-progress')
  el.classList.add('visible')
  el.textContent = text + (detail ? '\n' + detail : '')
}

function hideProgress() {
  const el = document.getElementById('ai-progress')
  el.classList.remove('visible')
}

function extractHTML(raw) {
  const fenceMatch = raw.match(/```html\s*([\s\S]*?)```/i)
  if (fenceMatch) return fenceMatch[1].trim()
  const genericFence = raw.match(/```\s*([\s\S]*?)```/)
  if (genericFence) return genericFence[1].trim()
  const startIdx = raw.indexOf('<!DOCTYPE')
  if (startIdx === -1) {
    const htmlIdx = raw.indexOf('<html')
    if (htmlIdx !== -1) return raw.slice(htmlIdx)
  }
  if (startIdx !== -1) return raw.slice(startIdx)
  return raw.trim()
}

function getFileIcon(type) {
  const icons = { pdf: '📕', docx: '📘', txt: '📄', md: '📝', json: '🔧', csv: '📊' }
  return icons[type] || '📄'
}

function formatSize(bytes) {
  if (!bytes) return '—'
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`
}

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function debounce(fn, delay) {
  let timer
  return (...args) => {
    clearTimeout(timer)
    timer = setTimeout(() => fn(...args), delay)
  }
}

// ── Panel HTML ───────────────────────────────────────────────────────────────

function renderAIPanel() {
  const body = document.getElementById('ai-panel-body')
  if (!body) return

  body.innerHTML = `
    <!-- Tabs -->
    <div class="ai-tabs">
      <button class="ai-tab active" id="ai-tab-generate">生成</button>
      <button class="ai-tab" id="ai-tab-modify">修改</button>
      <button class="ai-tab" id="ai-tab-memory">附件 <span class="memory-badge" id="memory-badge" style="display:none"></span></button>
      <button class="ai-tab" id="ai-tab-style">风格</button>
    </div>

    <!-- Generate Form -->
    <div id="ai-generate-form" class="ai-section" style="display:flex;flex-direction:column;gap:10px;">
      <div class="form-group">
        <label class="form-label">主题 / 标题 *</label>
        <input class="form-input" id="ai-topic" type="text" placeholder="如：人工智能发展趋势">
        <div class="intent-hint" id="intent-hint-generate"></div>
      </div>
      <div class="form-group">
        <label class="form-label">页数</label>
        <div class="input-row">
          <input class="form-input" id="ai-pages" type="number" value="8" min="3" max="30" style="width:80px;flex:none;"
          oninput="const h=document.getElementById('pages-hint');if(h)h.style.display=this.value>=20?'block':'none'">
          <select class="form-select" id="ai-lang">
            <option value="zh">中文</option>
            <option value="en">英文</option>
          </select>
        </div>
        <div id="pages-hint" style="display:none;font-size:11px;color:var(--text-muted);margin-top:4px;">
          ≥20 页需使用支持长输出的模型（如 DeepSeek-V3、GPT-4o 等）
        </div>
      </div>
      <button id="ai-generate-btn">✨ 生成 PPT</button>
    </div>

    <!-- Modify Form -->
    <div id="ai-modify-form" class="ai-section" style="display:none;flex-direction:column;gap:10px;">
      <div class="form-group">
        <label class="form-label">修改指令</label>
        <textarea class="form-textarea" id="ai-modify-instruction"
          placeholder="如：把标题改大，换成蓝色配色，添加数据可视化图表..."
          rows="4"></textarea>
        <div class="intent-hint" id="intent-hint-modify"></div>
      </div>
      <button id="ai-modify-btn">🔧 应用 AI 修改</button>
    </div>

    <!-- Memory Form -->
    <div id="ai-memory-form" class="ai-section" style="display:none;flex-direction:column;gap:10px;">
      <div id="memory-drop-zone" class="memory-drop-zone">
        <div class="memory-drop-icon">📁</div>
        <div class="memory-drop-text">拖拽文件到此处，或</div>
        <button id="memory-upload-btn" class="btn-secondary">选择文件</button>
        <div class="memory-drop-hint">支持 TXT · MD · DOCX · PDF · JSON · CSV</div>
      </div>
      <div id="memory-upload-progress" style="display:none;font-size:12px;color:var(--text-muted);"></div>
      <div id="memory-list"></div>
      <div class="memory-tip">
        💡 启用的文件会作为背景知识注入生成/修改 Prompt
      </div>
    </div>

    <!-- Style Form -->
    <div id="ai-style-form" class="ai-section" style="display:none;flex-direction:column;gap:12px;">
      <div class="form-group">
        <label class="form-label">风格模板</label>
        <select class="form-select" id="style-template-select"></select>
      </div>
      <div id="style-preview"></div>
      <div class="form-group">
        <label class="form-label">色温 <span id="style-color-temp-val"></span></label>
        <input type="range" id="style-color-temp" min="0" max="100" value="50" class="style-slider">
        <div class="slider-labels"><span>冷</span><span>暖</span></div>
      </div>
      <div class="form-group">
        <label class="form-label">对比度</label>
        <input type="range" id="style-contrast" min="0" max="100" value="50" class="style-slider">
        <div class="slider-labels"><span>低</span><span>高</span></div>
      </div>
      <div class="form-group">
        <label class="form-label">内容密度</label>
        <input type="range" id="style-density" min="0" max="100" value="50" class="style-slider">
        <div class="slider-labels"><span>简约</span><span>丰富</span></div>
      </div>
      <button id="style-save-btn" class="btn-secondary">保存风格设置</button>
    </div>

    <!-- Shared stop button + progress -->
    <button id="ai-stop-btn">⏹ 停止生成</button>
    <div id="ai-progress"></div>
  `
}
