# Slide X

一个强大的 HTML 演示文稿编辑器，支持 AI 生成、可视化编辑和多格式导出。

![Slide X 主界面](docs/screenshots/app-overview.png)

## ✨ 主要功能

- 🤖 **Slide AI**：内置 AI 助手，输入主题即可生成完整 PPT
- 📎 **附件上传**：上传文档（PDF/DOCX/TXT/MD）作为 AI 背景知识
- 🎨 **风格模板**：11 种预设配色方案，胶囊按钮 + 悬停预览卡片
- 🖼️ **图片提取风格**：上传 PPT 截图，AI 自动提取配色和布局风格
- ✏️ **智能修改**：AI 自动识别意图，支持单页智能修改
- 👁️ **可视化编辑**：双击元素直接编辑，所见即所得
- 📝 **源码编辑**：CodeMirror 代码编辑器，支持语法高亮
- 🔄 **撤销重做**：完整的历史记录，每个幻灯片独立管理
- 🎬 **全屏演示**：专业的演示模式，支持快捷键导航
- 📥 **导入 PPTX**：直接导入 .pptx 文件，自动解析幻灯片内容和图片
- 🗒️ **演讲者备注**：每页可添加备注，AI 自动生成，导出到 PPTX 备注栏
- 📦 **多格式导出**：PPTX（图片/可编辑）、PDF、PNG、JPEG
- 📐 **智能字号**：Canvas 精确测量，CJK 字符级换行，导出时自动防溢出

## 🚀 快速开始

### 安装

```bash
# 克隆项目
git clone https://github.com/WarriorXu0302/SlideGen.git
cd SlideGen

# 安装依赖
npm install

# 启动应用
npm start
```

### 首次使用

1. 点击右上角 ⚙️ 设置 API Key
2. 点击「✨ AI 生成」按钮
3. 输入主题，点击「生成 PPT」

### 基本工作流程

1. **配置 API Key** → 2. **输入主题生成** → 3. **AI 修改微调** → 4. **导出 PPT**

## 🤖 Slide AI 功能详解

### 生成 PPT（两步流程）

**第一步：生成大纲**
1. 点击工具栏「✨ AI 生成」按钮
2. 输入主题（如"人工智能发展趋势"）
3. 设置页数和语言
4. 点击「📝 生成大纲」

**第二步：确认并生成**
1. 预览生成的大纲，可直接编辑
2. 满意后点击「✨ 确认并生成 PPT」
3. Slide AI 将根据大纲生成完整的 HTML PPT

![AI 生成演示](docs/screenshots/ai-generate.png)

### 附件上传

上传文档让 AI 生成更专业、准确的内容：

1. 切换到「附件」标签页
2. 拖拽或点击上传文件（支持 TXT/MD/DOCX/PDF/JSON/CSV）
3. 勾选要使用的文件（可随时切换启用/禁用）
4. 生成时 AI 会自动注入已启用的文件内容作为背景知识

### 风格模板

1. 切换到「风格」标签页
2. 点击胶囊按钮选择预设模板（商务专业、极简主义、科技电光等 11 种）
3. 悬停查看风格预览卡片（配色、布局预览）
4. 调节色温、对比度、内容密度
5. 点击「保存风格设置」

![风格模板面板](docs/screenshots/style-templates.png)

**从图片提取风格：**
1. 点击「🖼️ 从图片提取风格」按钮
2. 上传一张 PPT 截图
3. AI 自动分析并提取配色方案和布局风格
4. 生成时将使用提取的风格

### 智能修改

1. 切换到「修改」标签页
2. 输入修改指令（如"把标题改大"、"换成蓝色配色"）
3. AI 会保持整体风格一致性进行修改

## 📖 编辑指南

**可视化编辑：**
- 双击任意文本元素直接编辑
- Enter 键确认，Escape 键取消
- 修改会自动同步到源码

**源码编辑：**
- 点击"源码"按钮打开代码编辑器
- CodeMirror 语法高亮
- 自动保存修改

**演讲者备注：**
- 底部备注栏，点击「📝 备注」展开/收起
- AI 生成时自动写入每页演讲要点
- 手动输入或编辑，实时保存
- 导出可编辑 PPTX 时自动写入备注栏

**撤销重做：**
- `Ctrl/Cmd + Z`：撤销
- `Ctrl/Cmd + Y` 或 `Ctrl/Cmd + Shift + Z`：重做
- 每个幻灯片独立记录历史

### 导出功能

| 格式 | 说明 |
|------|------|
| **PPTX** | 图片模式，像素完美，兼容 PowerPoint/WPS |
| **可编辑 PPTX** | 原生文本框，支持 CJK 字号自适应、图表/SVG 截图嵌入、演讲者备注 |
| **PDF** | 适合分享和打印 |
| **PNG/JPEG** | 单张或打包下载 |

**导出提示：** 可编辑 PPTX 导出完成后，如有文本缩放、图片失败等问题，会弹出详细警告提示。

![导出效果对比](docs/screenshots/export-result.png)

## ⌨️ 快捷键

| 功能 | Windows | macOS |
|------|---------|-------|
| 打开文件 | `Ctrl + O` | `Cmd + O` |
| 保存 | `Ctrl + S` | `Cmd + S` |
| 另存为 | `Ctrl + Shift + S` | `Cmd + Shift + S` |
| 撤销 | `Ctrl + Z` | `Cmd + Z` |
| 重做 | `Ctrl + Y` | `Cmd + Y` |
| 开始演示 | `F5` | `F5` |
| 退出演示 | `Esc` | `Esc` |
| 切换幻灯片 | `← →` / `↑ ↓` | `← →` / `↑ ↓` |

## ⚙️ API 配置

点击右上角齿轮图标进入设置：

- **API Base URL**：OpenAI 兼容接口地址（默认 `https://api.openai.com/v1`）
- **API Key**：你的 API 密钥
- **模型名称**：手动输入，或点击「查询」按钮自动拉取 `/models` 接口的可用模型列表

**模型参数：**

| 参数 | 说明 | 默认值 |
|------|------|--------|
| Max Tokens | 单次生成最大 token 数 | 16384 |
| Temperature | 生成随机性（0 = 确定，2 = 最随机） | 0.7 |
| Top P | 核采样概率（通常与 Temperature 二选一调整） | 1.0 |

支持 OpenAI、DeepSeek、本地 Ollama 等兼容接口。

## 🛠️ 开发

```bash
# 安装依赖
npm install

# 启动开发
npm start

# 打包 macOS
npm run build:mac

# 打包 Windows
npm run build:win

# 打包所有平台
npm run build:all
```

## 🙏 致谢

Slide X 的开发离不开以下优秀的开源项目：

- [Electron](https://www.electronjs.org/) - 跨平台桌面应用框架
- [electron-builder](https://www.electron.build/) - Electron 应用打包工具
- [CodeMirror 6](https://codemirror.net/) - 强大的代码编辑器组件
- [html2canvas](https://html2canvas.hertzen.com/) - HTML 转 Canvas 截图
- [jsPDF](https://github.com/parallax/jsPDF) - JavaScript PDF 生成库
- [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) - JavaScript PowerPoint 生成库
- [JSZip](https://stuk.github.io/jszip/) - JavaScript ZIP 压缩库
- [Mammoth](https://github.com/mwilliamson/mammoth.js) - DOCX 文档解析
- [pdf-parse](https://www.npmjs.com/package/pdf-parse) - PDF 文档解析
- [文多多 AiPPT](https://docmee.cn) - 商用级 AI 生成 PPT 解决方案
- [PPTist](https://github.com/pipipi-pikachu/PPTist) - 基于 Vue 的在线 PPT 编辑器
- [banana-slides](https://github.com/nicepkg/banana-slides) - AI 驱动的 PPT 生成应用
- [html2pptx](https://github.com/nicepkg/html2pptx) - HTML 转 PowerPoint 转换工具

---

**享受创作！** 🎨✨

## 📄 License

MIT License
