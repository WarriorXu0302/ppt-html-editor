/**
 * ai-panel.js — AI generation panel with Memory, Style, and Intent detection
 */

import { STYLE_TEMPLATES, getTemplateById, buildStylePrompt } from './style-templates.js'
import { detectIntent } from './intent-detector.js'

// ── Constants ────────────────────────────────────────────────────────────────

const SLIDE_WIDTH = 1280
const SLIDE_HEIGHT = 720

let currentAbortController = null
let onNewPPT = null
let onModifySlide = null
let getState = null

// Outline-first generation state
let currentOutline = ''
let generationPhase = 'idle' // 'idle' | 'outline' | 'ppt'

// Cached DOM elements for streaming (avoids repeated queries)
let cachedOutlineTextarea = null
let cachedOutlinePreview = null

// ── System Prompts ──────────────────────────────────────────────────────────

// ── Utility Functions ────────────────────────────────────────────────────────

function debounce(fn, delay) {
  let timer
  return (...args) => {
    clearTimeout(timer)
    timer = setTimeout(() => fn(...args), delay)
  }
}

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

/**
 * Remove markdown code fences from content
 */
function cleanMarkdownFences(text) {
  return text.replace(/```markdown\n?/g, '').replace(/```\n?/g, '')
}

/**
 * Debounced outline preview rendering (avoids O(n²) work during streaming)
 */
const debouncedRenderOutlinePreview = debounce((markdown) => {
  renderOutlinePreview(markdown)
}, 100)

// ── System Prompts ──────────────────────────────────────────────────────────

const SYSTEM_PROMPT_OUTLINE = `你是一个专业的 PPT 策划师。请根据用户主题生成 PPT 大纲。

输出格式（Markdown）：
# [PPT 标题]

## 第 1 页：封面
- 主标题
- 副标题

## 第 2 页：[页面标题]
- 要点 1
- 要点 2
- 要点 3

## 第 3 页：[页面标题]
...

要求：
- 每页 3-5 个要点，内容具体、有价值
- 内容逻辑清晰，层次分明
- 页面标题简洁有力
- 只输出大纲，不要任何解释`

const SYSTEM_PROMPT_FULL = `你是一个顶尖的 PPT 视觉设计师。你创作的每一个作品都像是花费无数小时精心打磨，由设计领域最顶尖的大师亲手制作。

## 设计哲学（核心理念）

你的作品不是"幻灯片"，而是视觉艺术品。每一页都应该：
- 像博物馆级海报一样精致
- 通过空间、形式、色彩传达信息，而非堆砌文字
- 展现"专家级工艺"——每个元素的位置、大小、颜色都经过深思熟虑

## 输出格式

严格遵守以下结构规范，每页用 <section data-slide="N" data-title="页面标题"> 包裹：

<!DOCTYPE html>
<html>
<head>
<style>
  body { margin: 0; padding: 0; }
  section[data-slide] {
    width: ${SLIDE_WIDTH}px;
    height: ${SLIDE_HEIGHT}px;
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

## 视觉设计规范（严格遵守）

### 配色哲学
- 选择与主题内容深度匹配的配色，展现"色彩作为信息系统"的理念
- 一种主色占 60-70% 视觉权重，1-2 种辅色，一种点睛强调色
- 封面和结语页使用深色背景，内容页可使用浅色背景形成"三明治"结构
- 色彩对比必须强烈有力，绝不使用 timid（怯懦）的平均分布配色

### 布局哲学
- 每页都是独特的构图，绝不重复相同排版
- 拥抱非对称、对角线流动、网格突破、大胆留白
- 可用布局：双栏（文字左/图示右）、图标+文字行、2x2或2x3网格卡片、半出血图片+内容叠加
- 数据展示：大号数字统计（60-72pt 数字+小标签）、对比栏、时间线/流程图

### 视觉元素
- 每页必须有视觉元素作为"锚点"——色块、几何图形、渐变、图标符号
- 使用 CSS 渐变创造深度和氛围：径向渐变、线性渐变、网格背景
- 可用技法：噪点纹理（通过 SVG filter）、几何图案、层叠透明度、戏剧性阴影
- 标题字号 44-52px，正文 16-20px，说明文字 12-14px

### 文字即设计
- 文字是视觉元素的一部分，不是信息的堆砌
- 每页文字极简：核心观点 1-3 个，绝不写长段落
- 列表项用图标或数字替代普通项目符号
- 大号关键词 > 小号解释文字，形成视觉层次

### 绝对禁止
- ❌ 标题下方画横线/装饰线（这是 AI 生成的明显特征）
- ❌ 纯文字页面（每页必须有视觉元素）
- ❌ 正文居中对齐（列表和段落必须左对齐）
- ❌ 所有页面使用相同布局（单调是业余的标志）
- ❌ 引用任何外部资源（无外链字体、图片用 CSS 渐变/几何图形代替）
- ❌ 低对比度文字（浅色背景上不用浅灰文字）
- ❌ 使用 Inter、Roboto、Arial 等通用字体描述（因为是 HTML，使用系统字体即可）
- ❌ 紫色渐变配白底（这是 AI 生成的陈词滥调）

### 技术要求
- 每页固定尺寸 ${SLIDE_WIDTH}×${SLIDE_HEIGHT}px，position: relative，overflow: hidden
- 使用内联 CSS，style 写在 <head> 或元素上
- 只输出完整 HTML 代码，不要任何解释文字
- 所有视觉效果用纯 CSS 实现：渐变、阴影、圆角、变换

## 大师级工艺检查

生成前自问：
1. 这看起来像是花了无数小时打磨的作品吗？
2. 每个元素的位置都是有意为之吗？
3. 色彩搭配会让人印象深刻吗？
4. 布局够大胆、够独特吗？

只���能回答"是"时，才输出最终作品。`

const SYSTEM_PROMPT_MODIFY = `你是一个专业的前端开发者，擅长修改 HTML PPT 幻灯片。
用户会给你一段幻灯片的 HTML 代码，以及修改指令。
请按指令修改 HTML，保持外层结构不变，宽度保持 ${SLIDE_WIDTH}px，高度保持 ${SLIDE_HEIGHT}px。
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

  // Generate Outline
  document.getElementById('ai-outline-btn').addEventListener('click', handleGenerateOutline)
  document.getElementById('ai-regenerate-outline-btn').addEventListener('click', handleGenerateOutline)
  document.getElementById('outline-edit-toggle').addEventListener('click', toggleOutlineEditMode)

  // Sync textarea changes to preview when exiting edit mode
  document.getElementById('ai-outline').addEventListener('blur', () => {
    renderOutlinePreview(document.getElementById('ai-outline').value)
  })

  // Generate PPT (from outline)
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

// ── Generate Outline ─────────────────────────────────────────────────────────

async function handleGenerateOutline() {
  // Guard against overlapping requests
  if (generationPhase !== 'idle') return

  const config = await window.electronAPI.getConfig()
  if (!config.apiKey) {
    alert('请先在设置中配置 API Key')
    openSettings()
    return
  }

  const topic = document.getElementById('ai-topic').value.trim()
  if (!topic) {
    document.getElementById('ai-topic').focus()
    return
  }

  const pages = Math.max(3, Math.min(30, parseInt(document.getElementById('ai-pages').value) || 8))
  const lang = document.getElementById('ai-lang').value

  generationPhase = 'outline'
  setGenerating(true)

  const memoryContent = await getSelectedMemoryContent()
  let userPrompt = `请为「${topic}」生成一个 ${pages} 页的 PPT 大纲。`
  if (lang === 'en') userPrompt += '\n请用英文生成。'
  else userPrompt += '\n请用中文生成。'

  if (memoryContent) {
    userPrompt += `\n\n参考以下背景知识：\n${memoryContent}`
  }

  showProgress('正在生成大纲...')

  // Clear outline and show section for streaming
  currentOutline = ''

  // Cache DOM elements for streaming callback
  cachedOutlineTextarea = document.getElementById('ai-outline')
  cachedOutlinePreview = document.getElementById('outline-preview')
  cachedOutlineTextarea.value = ''

  document.getElementById('outline-section').style.display = 'block'
  document.getElementById('ai-outline-btn').style.display = 'none'
  setOutlineEditMode(false)

  try {
    const outline = await streamCompletion(config, SYSTEM_PROMPT_OUTLINE, userPrompt, (chunk, total) => {
      // Real-time update outline
      currentOutline += chunk
      const cleanOutline = cleanMarkdownFences(currentOutline)
      cachedOutlineTextarea.value = cleanOutline
      // Use debounced preview to avoid O(n²) work
      debouncedRenderOutlinePreview(cleanOutline)
      showProgress(`正在生成大纲... ${total} 字符`)
    }, 128000)

    currentOutline = cleanMarkdownFences(outline).trim()
    cachedOutlineTextarea.value = currentOutline
    renderOutlinePreview(currentOutline)

    // Show outline section in preview mode
    document.getElementById('outline-section').style.display = 'block'
    document.getElementById('ai-outline-btn').style.display = 'none'
    setOutlineEditMode(false)

    showProgress('✓ 大纲生成完成！请确认后生成 PPT')
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
    generationPhase = 'idle'
    setGenerating(false)
    // Clear cached elements
    cachedOutlineTextarea = null
    cachedOutlinePreview = null
  }
}

// ── Generate PPT ─────────────────────────────────────────────────────────────

async function handleGenerate() {
  // Guard against overlapping requests
  if (generationPhase !== 'idle') return

  // Get outline from textarea (user may have edited it)
  const outline = document.getElementById('ai-outline')?.value?.trim()
  if (!outline) {
    alert('请先生成大纲')
    return
  }

  const config = await window.electronAPI.getConfig()
  if (!config.apiKey) {
    alert('请先在设置中配置 API Key')
    openSettings()
    return
  }

  generationPhase = 'ppt'
  setGenerating(true)

  const topic = document.getElementById('ai-topic').value.trim()
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

  // Build prompt with outline, memory and style
  const memoryContent = await getSelectedMemoryContent()
  const styleId = selectedStyleId
  const styleParams = getStyleParams()
  // Include extracted style description if available
  const extractedStyle = window._extractedStyleDesc || ''
  const userPrompt = buildGeneratePromptWithOutline(topic, pages, lang, outline, memoryContent, styleId, styleParams, extractedStyle)
  const systemPrompt = buildSystemPromptWithStyle(styleId)

  const maxTokens = 128000

  showProgress('正在根据大纲生成 PPT...')

  try {
    const html = await streamCompletion(config, systemPrompt, userPrompt, (chunk, total) => {
      showProgress(`正在生成... ${total} 字符`, chunk.slice(-200))
    }, maxTokens)

    const raw = extractHTML(html)
    const clean = validateAndFixSlides(raw)
    if (onNewPPT) onNewPPT(clean)
    showProgress('✓ 生成完成！', '')

    // Reset outline state for next generation
    resetOutlineState()

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
    generationPhase = 'idle'
    setGenerating(false)
  }
}

function resetOutlineState() {
  currentOutline = ''
  document.getElementById('ai-outline').value = ''
  document.getElementById('outline-preview').innerHTML = ''
  document.getElementById('outline-section').style.display = 'none'
  document.getElementById('ai-outline-btn').style.display = 'block'
  // Reset to preview mode
  setOutlineEditMode(false)
}

function toggleOutlineEditMode() {
  const textarea = document.getElementById('ai-outline')
  const preview = document.getElementById('outline-preview')
  const isEditing = textarea.style.display !== 'none'
  setOutlineEditMode(!isEditing)
}

function setOutlineEditMode(editing) {
  const textarea = document.getElementById('ai-outline')
  const preview = document.getElementById('outline-preview')
  const toggleBtn = document.getElementById('outline-edit-toggle')

  if (editing) {
    textarea.style.display = 'block'
    preview.style.display = 'none'
    toggleBtn.textContent = '👁️ 预览'
    textarea.focus()
  } else {
    textarea.style.display = 'none'
    preview.style.display = 'block'
    toggleBtn.textContent = '✏️ 编辑'
    renderOutlinePreview(textarea.value)
  }
}

function renderOutlinePreview(markdown) {
  const preview = document.getElementById('outline-preview')
  if (!preview || !markdown) {
    if (preview) preview.innerHTML = '<div class="outline-empty">暂无大纲</div>'
    return
  }

  // Simple markdown to HTML conversion
  let html = escapeHtml(markdown)
    // Headers
    .replace(/^# (.+)$/gm, '<h1>$1</h1>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    // List items
    .replace(/^- (.+)$/gm, '<li>$1</li>')
    // Bold
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    // Line breaks
    .replace(/\n/g, '')

  // Wrap consecutive <li> in <ul>
  html = html.replace(/(<li>.*?<\/li>)+/g, '<ul>$&</ul>')

  preview.innerHTML = html
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

function buildGeneratePromptWithOutline(topic, pages, lang, outline, memoryContent, styleId, styleParams, extractedStyle = '') {
  let prompt = `请严格按照以下大纲结构生成 PPT，主题「${topic}」，共 ${pages} 页。

## 大纲结构（必须严格遵循）

${outline}

## 生成要求

- 每个 "## 第 N 页" 对应一个 <section data-slide="N">
- 页面标题使用大纲中的标题
- 要点内容展开为具体的视觉化设计
- 必须生成从 data-slide="1" 到 data-slide="${pages}" 的全部页面`

  if (lang === 'en') prompt += '\n- 请用英文生成所有文字内容。'
  else prompt += '\n- 请用中文生成所有文字内容。'

  // Append style instructions
  if (styleId) {
    const styleHint = buildStylePrompt(styleId, styleParams)
    if (styleHint) prompt += styleHint
  } else if (extractedStyle) {
    // Use extracted style if no template selected
    prompt += `\n\n## 风格指令（从图片提取）\n\n${extractedStyle}\n\n请完全按照这个风格设计每一页。`
  }

  // Append memory content as background knowledge
  if (memoryContent) {
    prompt += `\n\n以下是相关背景知识，请据此生成更准确、专业的内容：\n\n${memoryContent}`
  }

  return prompt
}

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

  return SYSTEM_PROMPT_FULL + `

## 选定风格：${template.emoji} ${template.name}

**设计哲学**: ${template.designPhilosophy}

**调色板要求**（必须严格使用这些颜色）:
- 主色: ${template.colors.primary}
- 辅色: ${template.colors.secondary}
- 强调色: ${template.colors.accent}
- 背景色: ${template.colors.background}
- 主文字色: ${template.colors.text}
- 次要文字色: ${template.colors.textMuted}

**字体**: 标题 font-weight: ${template.fonts.title === 'bold' ? '700' : '400'}, 正文 font-weight: ${template.fonts.body === 'bold' ? '700' : template.fonts.body === 'light' ? '300' : '400'}

**布局风格**: ${template.layout}

请完全按照这个风格设计每一页，让整个 PPT 风格统一、专业、令人印象深刻。`
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

let selectedStyleId = ''
let stylePreviewPopup = null
let hidePopupTimeout = null

function initStylePanel() {
  const capsulesContainer = document.getElementById('style-capsules')
  if (!capsulesContainer) return

  // Render capsule buttons
  renderStyleCapsules(capsulesContainer)

  // Create popup element and attach to body for correct positioning
  createStylePreviewPopup()
  setupStyleCapsuleHover()

  // Setup sliders
  document.getElementById('style-color-temp').addEventListener('input', updateStylePreview)
  document.getElementById('style-contrast').addEventListener('input', updateStylePreview)
  document.getElementById('style-density').addEventListener('input', updateStylePreview)
  document.getElementById('style-save-btn').addEventListener('click', saveStyleConfig)

  // Setup image extraction
  const extractBtn = document.getElementById('style-extract-btn')
  const imageInput = document.getElementById('style-image-input')
  if (extractBtn && imageInput) {
    extractBtn.addEventListener('click', () => imageInput.click())
    imageInput.addEventListener('change', handleStyleImageExtract)
  }
}

function createStylePreviewPopup() {
  // Remove existing popup if any
  const existing = document.getElementById('style-preview-popup-dynamic')
  if (existing) existing.remove()

  // Create popup and attach to body
  const popup = document.createElement('div')
  popup.id = 'style-preview-popup-dynamic'
  popup.className = 'style-preview-popup'
  popup.innerHTML = `
    <div class="style-preview-popup-image" id="style-popup-image"></div>
    <div class="style-preview-popup-content">
      <div class="style-preview-popup-title" id="style-popup-title"></div>
      <div class="style-preview-popup-desc" id="style-popup-desc"></div>
      <div class="style-preview-popup-colors" id="style-popup-colors"></div>
    </div>
  `
  document.body.appendChild(popup)
  stylePreviewPopup = popup
}

function renderStyleCapsules(container) {
  // "None" capsule
  let html = `<button class="style-capsule style-capsule-none ${!selectedStyleId ? 'selected' : ''}"
    data-style-id="">✨ 自定义</button>`

  // Template capsules
  html += STYLE_TEMPLATES.map(t => `
    <button class="style-capsule ${selectedStyleId === t.id ? 'selected' : ''}"
      data-style-id="${t.id}">
      <span class="style-dot" style="background:${t.color}"></span>
      ${t.emoji} ${t.name}
    </button>
  `).join('')

  container.innerHTML = html

  // Add click handlers
  container.querySelectorAll('.style-capsule').forEach(btn => {
    btn.addEventListener('click', () => {
      selectedStyleId = btn.dataset.styleId || ''
      // Update selection visual
      container.querySelectorAll('.style-capsule').forEach(b => b.classList.remove('selected'))
      btn.classList.add('selected')
      updateStylePreview()
    })
  })
}

function setupStyleCapsuleHover() {
  const container = document.getElementById('style-capsules')
  if (!container || !stylePreviewPopup) return

  container.addEventListener('mouseover', (e) => {
    const capsule = e.target.closest('.style-capsule')
    if (!capsule) return

    const styleId = capsule.dataset.styleId
    if (!styleId) {
      hideStylePreviewPopup()
      return
    }

    clearTimeout(hidePopupTimeout)
    showStylePreviewPopup(styleId, capsule)
  })

  container.addEventListener('mouseleave', () => {
    hidePopupTimeout = setTimeout(hideStylePreviewPopup, 100)
  })
}

function showStylePreviewPopup(styleId, capsule) {
  const template = getTemplateById(styleId)
  if (!template || !stylePreviewPopup) return

  // Update popup content
  const imageEl = document.getElementById('style-popup-image')
  const titleEl = document.getElementById('style-popup-title')
  const descEl = document.getElementById('style-popup-desc')
  const colorsEl = document.getElementById('style-popup-colors')

  if (!imageEl || !titleEl || !descEl || !colorsEl) return

  // Render CSS-based preview (mini slide preview) - escape text content
  const safeName = escapeHtml(template.name)
  const safeDesc = escapeHtml(template.description)
  imageEl.innerHTML = `
    <div style="width:100%;height:100%;padding:12px;background:${template.colors.background};display:flex;flex-direction:column;justify-content:center;">
      <div style="font-size:16px;font-weight:bold;color:${template.colors.text};margin-bottom:6px;">${safeName}</div>
      <div style="font-size:10px;color:${template.colors.textMuted};margin-bottom:8px;">${safeDesc}</div>
      <div style="display:flex;gap:4px;">
        <div style="width:40px;height:4px;background:${template.colors.primary};border-radius:2px;"></div>
        <div style="width:20px;height:4px;background:${template.colors.accent};border-radius:2px;"></div>
      </div>
    </div>
  `
  imageEl.className = 'style-preview-popup-image'

  titleEl.innerHTML = `<span style="color:${template.color}">${template.emoji}</span> ${safeName}`
  descEl.textContent = template.designPhilosophy  // textContent is safe

  // Render color swatches
  const colors = [
    template.colors.primary,
    template.colors.secondary,
    template.colors.accent,
    template.colors.background
  ]
  colorsEl.innerHTML = colors.map(c =>
    `<div class="style-preview-popup-color" style="background:${escapeHtml(c)}"></div>`
  ).join('')

  // Make popup visible first to get actual dimensions
  stylePreviewPopup.style.visibility = 'hidden'
  stylePreviewPopup.classList.add('visible')

  // Get actual popup dimensions
  const popupRect = stylePreviewPopup.getBoundingClientRect()
  const popupWidth = popupRect.width || 260
  const popupHeight = popupRect.height || 220

  // Position popup above the capsule
  const rect = capsule.getBoundingClientRect()

  let left = rect.left + (rect.width / 2) - (popupWidth / 2)
  let top = rect.top - popupHeight - 8

  // Keep within viewport
  if (left < 10) left = 10
  if (left + popupWidth > window.innerWidth - 10) {
    left = window.innerWidth - popupWidth - 10
  }
  if (top < 10) {
    // Show below if no space above
    top = rect.bottom + 8
  }

  stylePreviewPopup.style.left = `${left}px`
  stylePreviewPopup.style.top = `${top}px`
  stylePreviewPopup.style.visibility = 'visible'
}

function hideStylePreviewPopup() {
  if (stylePreviewPopup) {
    stylePreviewPopup.classList.remove('visible')
    stylePreviewPopup.style.visibility = 'hidden'
  }
}

async function handleStyleImageExtract(e) {
  const file = e.target.files?.[0]
  if (!file) return

  const extractBtn = document.getElementById('style-extract-btn')
  const originalText = extractBtn.textContent
  extractBtn.textContent = '🔄 提取中...'
  extractBtn.disabled = true

  try {
    // Read image as base64
    const base64 = await readFileAsBase64(file)
    const mimeType = file.type || 'image/png'

    // Call AI API to extract style
    const config = await window.electronAPI.getConfig()
    if (!config.apiKey) {
      throw new Error('请先在设置中配置 API Key')
    }

    const response = await fetch(`${config.baseUrl || 'https://api.openai.com/v1'}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${config.apiKey}`
      },
      body: JSON.stringify({
        model: config.model || 'gpt-4o',
        messages: [{
          role: 'user',
          content: [
            {
              type: 'image_url',
              image_url: { url: `data:${mimeType};base64,${base64}` }
            },
            {
              type: 'text',
              text: `分析这张 PPT/演示文稿截图的视觉风格，提取以下信息：
1. 配色方案（主色、辅色、强调色、背景色、文字色，hex格式）
2. 布局风格（极简/现代/商务/创意/学术等）
3. 字体风格建议（粗细、大小）
4. 视觉特征（渐变/阴影/圆角/几何图形等）

以简洁的风格描述文本输出，不要 JSON 格式。示例：
"深色科技风格：纯黑背景(#0A0A0F)，霓虹蓝主色(#0066FF)配青色辅色(#00FFFF)，橙色强调。标题粗体大字号，高对比度。使用发光边框和网格线背景。"`
            }
          ]
        }],
        max_tokens: 500
      })
    })

    if (!response.ok) {
      const errData = await response.json().catch(() => ({}))
      throw new Error(errData.error?.message || `API 请求失败: ${response.status}`)
    }

    const data = await response.json()
    const styleDesc = data.choices?.[0]?.message?.content || ''

    if (styleDesc) {
      // Clear selected template and show extracted style
      selectedStyleId = ''
      renderStyleCapsules(document.getElementById('style-capsules'))

      // Show the extracted style description in the preview area
      const previewEl = document.getElementById('style-preview')
      if (previewEl) {
        previewEl.innerHTML = `
          <div class="style-preview-card" style="background:var(--bg-surface);border:1px solid var(--accent);padding:12px;">
            <div style="font-size:12px;color:var(--accent);margin-bottom:6px;">🖼️ 提取的风格</div>
            <div style="font-size:11px;color:var(--text-secondary);line-height:1.6;">${escapeHtml(styleDesc)}</div>
          </div>
        `
      }

      // Store extracted style for later use
      window._extractedStyleDesc = styleDesc
    }
  } catch (err) {
    console.error('Style extraction failed:', err)
    alert('风格提取失败: ' + err.message)
  } finally {
    extractBtn.textContent = originalText
    extractBtn.disabled = false
    e.target.value = ''
  }
}

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      const base64 = reader.result.split(',')[1]
      resolve(base64)
    }
    reader.onerror = reject
    reader.readAsDataURL(file)
  })
}

function applyConfigToForms(config) {
  // Settings form
  document.getElementById('settings-base-url').value = config.baseUrl || 'https://api.openai.com/v1'
  document.getElementById('settings-api-key').value = config.apiKey || ''
  document.getElementById('settings-model').value = config.model || 'gpt-4o'

  // Style form
  if (config.styleConfig) {
    if (config.styleConfig.templateId) {
      selectedStyleId = config.styleConfig.templateId
      const container = document.getElementById('style-capsules')
      if (container) renderStyleCapsules(container)
    }
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
  const templateId = selectedStyleId
  const previewEl = document.getElementById('style-preview')
  if (!previewEl) return

  // Don't override if we have an extracted style
  if (window._extractedStyleDesc && !templateId) {
    return
  }

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

  // Clear extracted style when a template is selected
  if (templateId) {
    window._extractedStyleDesc = null
  }
}

async function saveStyleConfig() {
  const templateId = selectedStyleId
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
  const outlineBtn = document.getElementById('ai-outline-btn')
  const regenerateBtn = document.getElementById('ai-regenerate-outline-btn')
  const genBtn = document.getElementById('ai-generate-btn')
  const modBtn = document.getElementById('ai-modify-btn')
  const stopBtn = document.getElementById('ai-stop-btn')

  if (isModify) {
    modBtn.disabled = active
  } else {
    if (outlineBtn) outlineBtn.disabled = active
    if (regenerateBtn) regenerateBtn.disabled = active
    if (genBtn) genBtn.disabled = active
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

/**
 * Post-process generated HTML to fix common issues:
 * - Ensure all slides have correct dimensions
 * - Fix missing position/overflow styles
 * - Add missing data-slide attributes
 */
function validateAndFixSlides(html) {
  // Parse HTML
  const parser = new DOMParser()
  const doc = parser.parseFromString(html, 'text/html')

  // Find all slide sections
  const slides = doc.querySelectorAll('section[data-slide], section.slide, .slide, [class*="slide"]')

  if (slides.length === 0) {
    // Try to find any section elements
    const sections = doc.querySelectorAll('section')
    sections.forEach((section, i) => {
      if (!section.hasAttribute('data-slide')) {
        section.setAttribute('data-slide', String(i + 1))
      }
      fixSlideStyles(section)
    })
  } else {
    slides.forEach((slide, i) => {
      if (!slide.hasAttribute('data-slide')) {
        slide.setAttribute('data-slide', String(i + 1))
      }
      fixSlideStyles(slide)
    })
  }

  // Ensure global styles exist
  let styleTag = doc.querySelector('style')
  if (!styleTag) {
    styleTag = doc.createElement('style')
    doc.head.appendChild(styleTag)
  }

  // Add/ensure base slide styles
  const baseStyles = `
section[data-slide] {
  width: ${SLIDE_WIDTH}px !important;
  height: ${SLIDE_HEIGHT}px !important;
  position: relative !important;
  overflow: hidden !important;
  box-sizing: border-box !important;
}
`
  if (!styleTag.textContent.includes(`${SLIDE_WIDTH}px`)) {
    styleTag.textContent = baseStyles + styleTag.textContent
  }

  // Serialize back to HTML
  return '<!DOCTYPE html>\n' + doc.documentElement.outerHTML
}

function fixSlideStyles(element) {
  // Ensure critical inline styles
  const style = element.style
  style.width = `${SLIDE_WIDTH}px`
  style.height = `${SLIDE_HEIGHT}px`
  style.position = 'relative'
  style.overflow = 'hidden'
  style.boxSizing = 'border-box'
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

      <!-- Step 1: Generate Outline -->
      <button id="ai-outline-btn" class="btn-outline-generate">📝 生成大纲</button>

      <!-- Outline Section (hidden until outline is generated) -->
      <div id="outline-section" class="outline-section" style="display:none;">
        <div class="outline-header">
          <label class="form-label">大纲预览</label>
          <button id="outline-edit-toggle" class="btn-text">✏️ 编辑</button>
        </div>
        <div id="outline-preview" class="outline-preview"></div>
        <textarea class="form-textarea outline-textarea" id="ai-outline" rows="10" style="display:none;" placeholder="大纲将在此显示..."></textarea>
        <div class="outline-actions">
          <button id="ai-regenerate-outline-btn" class="btn-secondary">🔄 重新生成</button>
          <button id="ai-generate-btn">✨ 确认并生成 PPT</button>
        </div>
      </div>
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
        <div class="style-capsules" id="style-capsules"></div>
      </div>
      <button class="style-extract-btn" id="style-extract-btn">
        🖼️ 从图片提取风格
      </button>
      <input type="file" id="style-image-input" accept="image/*" hidden>
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
