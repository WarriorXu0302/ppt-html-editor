/**
 * style-templates.js — PPT style templates
 * Inspired by theme-factory, canvas-design, and frontend-design skills
 *
 * Design Philosophy Principles:
 * - Each theme is a visual philosophy/aesthetic movement
 * - Colors must have strong intentionality, not generic AI defaults
 * - Typography choices should be distinctive and characterful
 * - Layouts should feel "meticulously crafted by a master designer"
 */

export const STYLE_TEMPLATES = [
  // ─── Professional / Business ───────────────────────────────────────────────
  {
    id: 'business',
    name: '商务专业',
    emoji: '💼',
    description: '深蓝系，沉稳有力',
    color: '#1E2761',  // main theme color for UI indicator
    previewImage: './presets/business.png',
    colors: {
      primary: '#1E2761',    // navy
      secondary: '#CADCFC',  // ice blue
      accent: '#F59E0B',     // amber
      background: '#0F1728',
      text: '#F0F4FF',
      textMuted: '#8A9BB5'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '48px', bodySize: '19px' },
    layout: 'formal',
    designPhilosophy: 'Concrete Poetry - 通过纪念碑式的形式和大胆几何传达信息',
    promptHint: '商务深蓝风格，navy蓝主色配冰蓝辅色，琥珀色强调，沉稳专业，封面深色，内容页可浅色交替。视觉元素：几何色块、数据可视化图表、大号数字统计。布局严谨但不呆板，通过空间张力传达专业感。'
  },
  {
    id: 'clean',
    name: '极简主义',
    emoji: '✨',
    description: '白底，炭灰+深蓝',
    color: '#36454F',
    previewImage: './presets/clean.png',
    colors: {
      primary: '#36454F',    // charcoal
      secondary: '#708090',  // slate
      accent: '#0066CC',
      background: '#FFFFFF',
      text: '#1A1A2E',
      textMuted: '#6B7280'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '44px', bodySize: '18px' },
    layout: 'minimal',
    designPhilosophy: 'Geometric Silence - 纯粹的秩序与克制，网格精确但留白大胆',
    promptHint: '极简清新风格，纯白背景，炭灰和石板色为主。大量留白(每页至少40%空白)，单一视觉焦点。禁止装饰线、禁止渐变背景。图形极简：细线框、单色图标、纯色色块。瑞士风格排版。'
  },
  {
    id: 'academic',
    name: '学术报告',
    emoji: '📚',
    description: '深蓝严谨，数据友好',
    color: '#065A82',
    previewImage: './presets/academic.png',
    colors: {
      primary: '#065A82',    // deep blue
      secondary: '#1C7293',  // teal
      accent: '#21295C',     // midnight
      background: '#F8FAFF',
      text: '#1A2535',
      textMuted: '#5A6B7A'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '40px', bodySize: '17px' },
    layout: 'academic',
    designPhilosophy: 'Systematic Observation - 将抽象概念以科学图表的视觉语言呈现',
    promptHint: '学术报告风格，深海蓝色系，浅白背景。强调数据可视化：柱状图、折线图、流程图、对比表格。章节标题清晰，层级分明。正文左对齐，引用和注释用小字。参考文献格式规范。'
  },

  // ─── Technology / Innovation ───────────────────────────────────────────────
  {
    id: 'tech',
    name: '科技电光',
    emoji: '⚡',
    description: '纯黑底，电光蓝青',
    color: '#0066FF',
    previewImage: './presets/tech.png',
    colors: {
      primary: '#0066FF',    // electric blue
      secondary: '#00FFFF',  // neon cyan
      accent: '#FF6B35',
      background: '#0A0A0F',
      text: '#FFFFFF',
      textMuted: '#9DA8B7'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '52px', bodySize: '18px' },
    layout: 'tech',
    designPhilosophy: 'Tech Innovation - 大胆现代的科技美学，高对比度霓虹与深邃黑',
    promptHint: '科技感风格，纯黑背景(#0A0A0F)，电光蓝和霓虹青配色。高对比度视觉冲击。使用：发光边框、渐变描边、网格线背景、数据流动效果、圆角卡片。禁止使用标题装饰线。图标风格：线性、发光。'
  },
  {
    id: 'midnight-galaxy',
    name: '午夜星系',
    emoji: '🌌',
    description: '深紫+宇宙蓝，戏剧感',
    color: '#6366F1',
    previewImage: './presets/midnight-galaxy.png',
    colors: {
      primary: '#6366F1',    // indigo
      secondary: '#A78BFA',  // violet
      accent: '#F472B6',     // pink
      background: '#0F0A1A',
      text: '#E8E4F0',
      textMuted: '#8A7BAA'
    },
    fonts: { title: 'bold', body: 'light', titleSize: '50px', bodySize: '19px' },
    layout: 'dark-elegant',
    designPhilosophy: 'Midnight Galaxy - 宇宙深邃的戏剧性，渐变与星尘的诗意',
    promptHint: '宇宙星系风格，深紫/深蓝渐变背景。使用：径向渐变、光晕效果、粒子点缀、流体形状。色彩过渡柔和。银色/薰衣草色文字。适合创意、娱乐、AI/未来主题。每页可有微妙的星空或极光效果。'
  },

  // ─── Creative / Energetic ──────────────────────────────────────────────────
  {
    id: 'coral-energy',
    name: '活力珊瑚',
    emoji: '🔥',
    description: '珊瑚红+金色，充满活力',
    color: '#F96167',
    previewImage: './presets/coral-energy.png',
    colors: {
      primary: '#F96167',    // coral
      secondary: '#F9E795',  // gold
      accent: '#2F3C7E',     // navy
      background: '#1A1025',
      text: '#FFF5F5',
      textMuted: '#D4A0A0'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '50px', bodySize: '19px' },
    layout: 'creative',
    designPhilosophy: 'Chromatic Language - 色彩作为主要信息系统，鲜艳大胆',
    promptHint: '活力珊瑚风格，珊瑚红主色、金色辅色、海军蓝点缀。鲜艳对比强烈。使用：大色块分割、几何图形叠加、动感斜线、emoji图标。适合营销、发布会、创意pitch。布局打破常规，可不对称。'
  },
  {
    id: 'golden-hour',
    name: '金色时刻',
    emoji: '🌅',
    description: '日落暖金，奢华温暖',
    color: '#D4A574',
    previewImage: './presets/golden-hour.png',
    colors: {
      primary: '#D4A574',    // gold
      secondary: '#E8C39E',  // champagne
      accent: '#8B4513',     // saddle brown
      background: '#1A150F',
      text: '#FFF8F0',
      textMuted: '#C9B8A8'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '48px', bodySize: '18px' },
    layout: 'luxury',
    designPhilosophy: 'Golden Hour - 温暖奢华的日落美学，黄金比例与柔光',
    promptHint: '金色时刻风格，奢华暖金色调。深色背景配金色/香槟色元素。使用：渐变金属质感、柔和光晕、优雅曲线、细线装饰。适合高端品牌、金融、奢侈品。文字可用衬线体风格。'
  },

  // ─── Nature / Organic ──────────────────────────────────────────────────────
  {
    id: 'teal-trust',
    name: '蓝绿信任',
    emoji: '🌊',
    description: '青蓝系，清爽专业',
    color: '#0891B2',
    previewImage: './presets/teal-trust.png',
    colors: {
      primary: '#0891B2',    // teal
      secondary: '#06B6D4',  // cyan
      accent: '#10B981',     // emerald
      background: '#F0FDFA',
      text: '#134E4A',
      textMuted: '#5E8A87'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '44px', bodySize: '18px' },
    layout: 'clean',
    designPhilosophy: 'Ocean Depths - 海洋的深邃与清澈，专业而令人信任',
    promptHint: '蓝绿信任风格，浅白/薄荷背景，青蓝渐变色系。清爽专业、可信赖感。使用：圆角卡片、柔和阴影、波浪线条、水滴/叶片图形。适合医疗、环保、科技、教育。图表清晰易读。'
  },
  {
    id: 'terracotta',
    name: '暖砖大地',
    emoji: '🌵',
    description: '赤土色+沙色，温暖自然',
    color: '#B85042',
    previewImage: './presets/terracotta.png',
    colors: {
      primary: '#B85042',    // terracotta
      secondary: '#E7E8D1',  // sand
      accent: '#A7BEAE',     // sage
      background: '#2A1A17',
      text: '#F5EDE8',
      textMuted: '#C4A89A'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '46px', bodySize: '18px' },
    layout: 'warm',
    designPhilosophy: 'Desert Rose - 大地的温暖与沙漠的柔和，自然有机',
    promptHint: '暖砖大地风格，赤土主色、沙色辅色、鼠尾草绿点缀。温暖自然感，有机形状。使用：手绘风线条、不规则形状、纹理背景、植物/自然元素。适合文化、旅游、食品、手工艺。'
  },
  {
    id: 'botanical',
    name: '植物花园',
    emoji: '🌿',
    description: '森林绿+奶油白',
    color: '#2D5A27',
    previewImage: './presets/botanical.png',
    colors: {
      primary: '#2D5A27',    // forest green
      secondary: '#6B8E23',  // olive
      accent: '#F0E68C',     // khaki
      background: '#FFFEF7',
      text: '#1A2E1A',
      textMuted: '#5A6B5A'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '44px', bodySize: '18px' },
    layout: 'organic',
    designPhilosophy: 'Botanical Garden - 新鲜有机的花园色彩，生命力与平衡',
    promptHint: '植物花园风格，森林绿主色、奶油白背景。新鲜自然、有机平衡。使用：叶片形状、自然纹理、手绘植物图案、柔和曲线。布局呼吸感强。适合环保、健康、农业、自然教育。'
  },

  // ─── Arctic / Cool ─────────────────────────────────────────────────────────
  {
    id: 'arctic-frost',
    name: '极地霜冻',
    emoji: '❄️',
    description: '冰蓝+银白，冷冽清新',
    color: '#38BDF8',
    previewImage: './presets/arctic-frost.png',
    colors: {
      primary: '#38BDF8',    // sky blue
      secondary: '#7DD3FC',  // light sky
      accent: '#0EA5E9',     // azure
      background: '#F8FAFC',
      text: '#0F172A',
      textMuted: '#64748B'
    },
    fonts: { title: 'bold', body: 'normal', titleSize: '46px', bodySize: '18px' },
    layout: 'crisp',
    designPhilosophy: 'Arctic Frost - 冬季的冷冽清新，冰晶般的精确与纯净',
    promptHint: '极地霜冻风格，冰蓝/天蓝主色、银白背景。冷冽清新、高清晰度。使用：锐利线条、几何冰晶图案、高光效果、玻璃质感。适合科技、AI、数据、医疗。强调精确与专业。'
  }
]

export function getTemplateById(id) {
  return STYLE_TEMPLATES.find(t => t.id === id) || null
}

/**
 * Build style prompt with design philosophy principles
 * Inspired by canvas-design and frontend-design skills
 */
export function buildStylePrompt(templateId, customParams = {}) {
  const template = getTemplateById(templateId)
  if (!template) return ''

  let prompt = `\n\n## 风格指令

**设计哲学**: ${template.designPhilosophy}

**具体要求**: ${template.promptHint}

**调色板** (必须严格使用):
- 主色: ${template.colors.primary}
- 辅色: ${template.colors.secondary}
- 强调色: ${template.colors.accent}
- 背景色: ${template.colors.background}
- 主文字: ${template.colors.text}
- 次要文字: ${template.colors.textMuted}

**字体风格**: 标题 ${template.fonts.title} ${template.fonts.titleSize}，正文 ${template.fonts.body} ${template.fonts.bodySize}`

  // Apply custom parameters
  if (customParams.colorTemp !== undefined) {
    const v = customParams.colorTemp
    if (v < 35) prompt += `\n- 整体偏冷色调处理，蓝/青/银色倾向`
    else if (v > 65) prompt += `\n- 整体偏暖色调处理，橙/金/红色倾向`
  }
  if (customParams.contrast !== undefined) {
    const v = customParams.contrast
    if (v > 70) prompt += `\n- 高对比度设计，文字清晰锐利，色块边界分明`
    else if (v < 30) prompt += `\n- 柔和低对比度，颜色过渡平滑，整体温和`
  }
  if (customParams.density !== undefined) {
    const v = customParams.density
    if (v > 70) prompt += `\n- 内容丰富密集，多用图表、数据、图标，信息量大`
    else if (v < 30) prompt += `\n- 极简留白风格，每页核心要素不超过3个，大量呼吸空间`
  }

  return prompt
}
