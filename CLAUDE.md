# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Slide X** is an Electron desktop application for editing and generating HTML-based presentations with AI assistance. The UI is in Chinese (zh-CN). The AI backend uses any OpenAI-compatible API (supports OpenAI, DeepSeek, local Ollama, etc.).

## Commands

```bash
npm install         # Install dependencies
npm start           # Run the Electron app in development
npm run build:mac   # Build for macOS (dmg, x64 + arm64)
npm run build:win   # Build for Windows (zip, x64)
npm run build:all   # Build for all platforms
```

No test framework is configured.

## Architecture

### Process Model (Electron)

- **main.js**: Main process — window management, native dialogs, IPC handlers, config storage (JSON in `userData/config.json`), file parsing (mammoth for DOCX, pdf-parse for PDF). Also handles `open-pptx-file` which reads PPTX bytes and returns them as a plain byte array to the renderer.
- **preload.js**: Secure IPC bridge exposing `window.electronAPI` to renderer. Tracks registered listeners in a `Map` to prevent duplicates.
- **renderer/**: All browser-context code (ES modules)

### Renderer Modules

- **app.js**: Main controller — state management (`window.appState`), navigation, undo/redo (per-slide history), presentation mode, welcome screen, speaker notes panel
- **parser.js**: Multi-strategy HTML parser supporting 4 formats:
  - `section-data-slide`: `<section data-slide="N">` elements
  - `iframe-srcdoc`: `<iframe srcdoc="...">` elements (preserves wrapper template)
  - `generic`: `<section>`, `div.slide`, or `div[class*="slide"]`
  - `raw`: Entire document as single slide
- **editor.js**: CodeMirror 6 source editor (loaded dynamically from esm.sh CDN) + visual inline editing via `contentEditable`
- **ai-panel.js**: AI generation panel with two-step flow (outline → PPT), streaming API, memory file context injection, style templates, intent detection, image-based style extraction, and single-slide modification
- **exporter.js**: Export to PPTX (image-based), editable PPTX (text extraction with CJK font sizing, chart/SVG screenshot embedding, speaker notes), PDF, PNG/JPEG
- **pptx-importer.js**: PPTX → HTML converter. Uses JSZip + DOMParser to parse `.pptx` XML, resolves slide layouts via a cache, and outputs slide objects `{ content, rawHtml, title, notes }` where `content` is a 1280×720px standalone HTML document.
- **style-templates.js**: Color palette definitions for 11 built-in style presets
- **intent-detector.js**: Keyword/pattern matching to route user input between generate vs. modify
- **styles.css**: All renderer CSS
- **index.html**: App shell — loads CDN UMD bundles, then `<script type="module" src="app.js">`
- **presets/**: Directory for user-saved style presets (currently empty, tracked with `.gitkeep`)

### Key IPC Channels

Invocable (`ipcRenderer.invoke` / `ipcMain.handle`):
- `read-file`, `write-file` — file I/O
- `show-open-dialog`, `show-save-dialog`, `show-message-box` — native dialogs
- `get-config`, `set-config` — config persistence
- `get-path` — Electron app paths
- `open-pptx-file` — opens native file picker and returns PPTX as byte array
- `open-presentation` — opens a separate fullscreen presentation window
- `show-memory-file-dialog`, `parse-memory-file`, `get-memory-list`, `save-memory-file`, `delete-memory-file`, `update-memory-tags` — memory file system

One-way (`ipcRenderer.send` / `ipcMain.on`):
- `set-title`, `set-document-edited`, `set-dirty-flag`, `save-complete`

Menu events sent from main to renderer: `menu-open`, `menu-save`, `menu-save-as`, `menu-undo`, `menu-redo`, `menu-open-file`, `menu-new`

### Key Design Patterns

1. **Slide Storage**: Each slide is a standalone HTML document (1280×720px) rendered in sandboxed iframes. Thumbnails use blob URLs with `IntersectionObserver` for lazy loading.

2. **HTML Reconstruction**: `reconstructHTML()` in parser.js rebuilds the full document from slides based on the detected format, preserving original wrapper templates for `iframe-srcdoc` format.

3. **State Flow**: State changes go through `applySlideContent()` which handles undo stack, dirty flag, thumbnail refresh, and preview updates.

4. **AI Two-Step Generation**: User first generates an outline (plain text), edits it, then confirms to generate the full HTML PPT. Prompts include system rules for slide design (CSS requirements, forbidden patterns like title underlines).

5. **AI Streaming**: Uses `fetch` with streaming `ReadableStream` body reader. Partial HTML chunks are accumulated and applied as they arrive.

6. **Memory Files**: AI context injection system — users attach documents (TXT/MD/DOCX/PDF/JSON/CSV) that are parsed and injected into AI prompts as background knowledge. Stored in `userData/memory/`, indexed in `config.json`.

7. **PPTX Import Flow**: `open-pptx-file` IPC → main.js reads file bytes → renderer receives byte array → `pptx-importer.js` uses JSZip to unpack XML → renders shapes, images, text as positioned HTML elements.

8. **Config Persistence**: Stored in Electron's `userData` as `config.json` — includes API settings (base URL, key, model, temperature, max tokens), recent files, memory files index, and style preferences.

### CDN Dependencies (loaded in index.html)

UMD bundles from jsDelivr (loaded before module scripts so they're available as globals):
- `html2canvas@1.4.1`, `jspdf@2.5.1`, `jszip@3.10.1`, `pptxgenjs@3.12.0`

CodeMirror 6 is loaded dynamically from `esm.sh` inside editor.js on first use.
