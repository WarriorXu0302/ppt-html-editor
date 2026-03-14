const { app, BrowserWindow, ipcMain, Menu, dialog, safeStorage, protocol } = require('electron')
const path = require('path')
const fs = require('fs')
const fsPromises = require('fs').promises
const os = require('os')
const { randomUUID } = require('crypto')

// ──── Constants ────
const MAX_MEMORY_FILE_CHARS = 50000
const MAX_RECENT_FILES = 10
const ALLOWED_HTML_EXTENSIONS = ['.html', '.htm']

let mainWindow
let isDirtyFlag = false
let config = {}
let pendingSaveResolve = null

const CONFIG_PATH = () => path.join(app.getPath('userData'), 'config.json')

// ──── Safe Storage Helpers ────
function encryptApiKey(apiKey) {
  if (!apiKey || !safeStorage.isEncryptionAvailable()) {
    return apiKey
  }
  try {
    return safeStorage.encryptString(apiKey).toString('base64')
  } catch (e) {
    console.error('Failed to encrypt API key:', e)
    return apiKey
  }
}

function decryptApiKey(encryptedKey) {
  if (!encryptedKey || !safeStorage.isEncryptionAvailable()) {
    return encryptedKey
  }
  try {
    // Check if it looks like base64 encrypted data
    if (encryptedKey.length > 100 && !encryptedKey.startsWith('sk-')) {
      const buffer = Buffer.from(encryptedKey, 'base64')
      return safeStorage.decryptString(buffer)
    }
    return encryptedKey
  } catch (e) {
    // Return as-is if decryption fails (might be plaintext from old version)
    return encryptedKey
  }
}

// ──── Config Management ────
async function loadConfig() {
  try {
    const data = await fsPromises.readFile(CONFIG_PATH(), 'utf8')
    config = JSON.parse(data)
    // Decrypt API key on load
    if (config.apiKey) {
      config.apiKey = decryptApiKey(config.apiKey)
    }
  } catch (e) {
    config = {
      recentFiles: [],
      apiKey: '',
      baseUrl: 'https://api.openai.com/v1',
      model: 'gpt-4o',
      memoryFiles: [],
      styleConfig: null
    }
  }
  // Ensure new fields exist on old configs
  if (!config.memoryFiles) config.memoryFiles = []
  if (config.styleConfig === undefined) config.styleConfig = null
  if (config.maxTokens === undefined) config.maxTokens = 16384
  if (config.temperature === undefined) config.temperature = 0.7
  if (config.topP === undefined) config.topP = 1.0
}

async function saveConfig() {
  try {
    await fsPromises.mkdir(path.dirname(CONFIG_PATH()), { recursive: true })
    // Create a copy with encrypted API key
    const configToSave = { ...config }
    if (configToSave.apiKey) {
      configToSave.apiKey = encryptApiKey(configToSave.apiKey)
    }
    await fsPromises.writeFile(CONFIG_PATH(), JSON.stringify(configToSave, null, 2), 'utf8')
  } catch (e) {
    console.error('Failed to save config:', e)
  }
}

function addRecentFile(filePath) {
  if (!config.recentFiles) config.recentFiles = []
  config.recentFiles = config.recentFiles.filter(f => f !== filePath)
  config.recentFiles.unshift(filePath)
  config.recentFiles = config.recentFiles.slice(0, MAX_RECENT_FILES)
  saveConfig()
  buildMenu()
}

// ──── Path Validation ────
function isValidFilePath(filePath, allowedExtensions = null) {
  if (!filePath || typeof filePath !== 'string') {
    return false
  }

  // Normalize and resolve the path
  const normalizedPath = path.normalize(filePath)
  const resolvedPath = path.resolve(filePath)

  // Check for path traversal attempts
  if (normalizedPath.includes('..') && resolvedPath !== normalizedPath) {
    return false
  }

  // Check extension if specified
  if (allowedExtensions) {
    const ext = path.extname(filePath).toLowerCase()
    if (!allowedExtensions.includes(ext)) {
      return false
    }
  }

  return true
}

function isAllowedReadPath(filePath) {
  // Allow reading from common safe locations
  const allowedRoots = [
    app.getPath('home'),
    app.getPath('documents'),
    app.getPath('downloads'),
    app.getPath('desktop'),
    app.getPath('userData'),
    os.tmpdir()
  ]

  const resolvedPath = path.resolve(filePath)
  return allowedRoots.some(root => resolvedPath.startsWith(root))
}

// ──── Menu Building ────
function buildRecentFilesMenu() {
  const recent = config.recentFiles || []
  if (recent.length === 0) {
    return [{ label: '(无最近文件)', enabled: false }]
  }
  return recent.map(filePath => ({
    label: path.basename(filePath),
    sublabel: filePath,
    click: () => mainWindow && mainWindow.webContents.send('menu-open-file', filePath)
  }))
}

function buildMenu() {
  const isMac = process.platform === 'darwin'

  const template = [
    ...(isMac ? [{
      label: app.name,
      submenu: [
        { role: 'about' },
        { type: 'separator' },
        { role: 'services' },
        { type: 'separator' },
        { role: 'hide' },
        { role: 'hideOthers' },
        { role: 'unhide' },
        { type: 'separator' },
        { role: 'quit' }
      ]
    }] : []),
    {
      label: '文件',
      submenu: [
        {
          label: '新建',
          accelerator: 'CmdOrCtrl+N',
          click: () => mainWindow && mainWindow.webContents.send('menu-new')
        },
        {
          label: '打开...',
          accelerator: 'CmdOrCtrl+O',
          click: () => mainWindow && mainWindow.webContents.send('menu-open')
        },
        { type: 'separator' },
        {
          label: '保存',
          accelerator: 'CmdOrCtrl+S',
          click: () => mainWindow && mainWindow.webContents.send('menu-save')
        },
        {
          label: '另存为...',
          accelerator: 'CmdOrCtrl+Shift+S',
          click: () => mainWindow && mainWindow.webContents.send('menu-save-as')
        },
        { type: 'separator' },
        {
          label: '最近打开',
          submenu: buildRecentFilesMenu()
        },
        { type: 'separator' },
        ...(isMac ? [] : [{ role: 'quit', label: '退出' }])
      ]
    },
    {
      label: '编辑',
      submenu: [
        {
          label: '撤销',
          accelerator: 'CmdOrCtrl+Z',
          click: () => mainWindow && mainWindow.webContents.send('menu-undo')
        },
        {
          label: '重做',
          accelerator: 'CmdOrCtrl+Y',
          click: () => mainWindow && mainWindow.webContents.send('menu-redo')
        },
        { type: 'separator' },
        { role: 'cut', label: '剪切' },
        { role: 'copy', label: '复制' },
        { role: 'paste', label: '粘贴' },
        { role: 'selectAll', label: '全选' }
      ]
    },
    {
      label: '视图',
      submenu: [
        { role: 'reload', label: '重新加载' },
        { role: 'toggleDevTools', label: '开发者工具' },
        { type: 'separator' },
        { role: 'resetZoom', label: '实际大小' },
        { role: 'zoomIn', label: '放大' },
        { role: 'zoomOut', label: '缩小' },
        { type: 'separator' },
        { role: 'togglefullscreen', label: '全屏' }
      ]
    },
    {
      label: '帮助',
      submenu: [
        {
          label: '关于 Slide X',
          click: async () => {
            await dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: '关于',
              message: 'Slide X',
              detail: 'Version 1.0.0\nAI-powered presentation editor'
            })
          }
        }
      ]
    }
  ]

  const menu = Menu.buildFromTemplate(template)
  Menu.setApplicationMenu(menu)
}

// ──── Window Creation ────
function createWindow() {
  const isMac = process.platform === 'darwin'

  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 900,
    minHeight: 600,
    titleBarStyle: isMac ? 'hiddenInset' : 'default',
    trafficLightPosition: isMac ? { x: 16, y: 16 } : undefined,
    acceptFirstMouse: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      // Note: webSecurity is disabled to allow loading CDN resources from file:// context
      // In production, consider using a custom protocol or local bundling
      webSecurity: false
    }
  })

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'))

  mainWindow.on('close', async (e) => {
    if (isDirtyFlag) {
      e.preventDefault()
      const choice = await dialog.showMessageBox(mainWindow, {
        type: 'question',
        buttons: ['保存', '不保存', '取消'],
        defaultId: 0,
        cancelId: 2,
        title: '未保存的更改',
        message: '文件已修改，是否保存？'
      })
      if (choice.response === 0) {
        // Request save and wait for confirmation
        const savePromise = new Promise(resolve => {
          pendingSaveResolve = resolve
          // Fallback timeout in case save-complete is never received
          setTimeout(() => {
            if (pendingSaveResolve) {
              pendingSaveResolve()
              pendingSaveResolve = null
            }
          }, 5000)
        })
        mainWindow.webContents.send('menu-save')
        await savePromise
        mainWindow.destroy()
      } else if (choice.response === 1) {
        mainWindow.destroy()
      }
      // response === 2: cancel, do nothing
    }
  })
}

// ──── IPC Handlers ────

ipcMain.handle('read-file', async (event, filePath) => {
  if (!isValidFilePath(filePath, ALLOWED_HTML_EXTENSIONS)) {
    throw new Error('Invalid file path or extension')
  }
  if (!isAllowedReadPath(filePath)) {
    throw new Error('Access to this path is not allowed')
  }
  return fsPromises.readFile(filePath, 'utf8')
})

ipcMain.handle('write-file', async (event, filePath, content) => {
  if (!isValidFilePath(filePath, ALLOWED_HTML_EXTENSIONS)) {
    throw new Error('Invalid file path or extension')
  }
  await fsPromises.mkdir(path.dirname(filePath), { recursive: true })
  await fsPromises.writeFile(filePath, content, 'utf8')
  addRecentFile(filePath)
  return true
})

ipcMain.handle('show-open-dialog', async () => {
  return dialog.showOpenDialog(mainWindow, {
    filters: [{ name: 'HTML Files', extensions: ['html', 'htm'] }],
    properties: ['openFile']
  })
})

ipcMain.handle('show-save-dialog', async (event, defaultPath) => {
  return dialog.showSaveDialog(mainWindow, {
    defaultPath: defaultPath || 'presentation.html',
    filters: [{ name: 'HTML Files', extensions: ['html', 'htm'] }]
  })
})

ipcMain.handle('get-config', () => config)

ipcMain.handle('set-config', (event, updates) => {
  config = { ...config, ...updates }
  saveConfig()
  if (updates.recentFiles !== undefined) buildMenu()
  return true
})

ipcMain.handle('show-message-box', async (event, options) => {
  return dialog.showMessageBox(mainWindow, options)
})

ipcMain.handle('get-path', (event, name) => {
  return app.getPath(name)
})

ipcMain.handle('open-pptx-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    filters: [{ name: 'PowerPoint Files', extensions: ['pptx'] }],
    properties: ['openFile']
  })
  if (result.canceled || !result.filePaths.length) return null
  const filePath = result.filePaths[0]
  const buffer = await fsPromises.readFile(filePath)
  // Transfer as plain array so it can pass through IPC serialization
  return { filePath, data: Array.from(buffer) }
})

ipcMain.on('set-title', (event, title) => {
  if (mainWindow) mainWindow.setTitle(title)
})

ipcMain.on('set-document-edited', (event, edited) => {
  if (mainWindow && process.platform === 'darwin') {
    mainWindow.setDocumentEdited(edited)
  }
})

ipcMain.on('set-dirty-flag', (event, dirty) => {
  isDirtyFlag = dirty
})

// Save completion notification for window close handling
ipcMain.on('save-complete', () => {
  if (pendingSaveResolve) {
    pendingSaveResolve()
    pendingSaveResolve = null
  }
})

// ──── Memory File Handlers ────

const ALLOWED_MEMORY_EXTENSIONS = ['txt', 'md', 'json', 'csv', 'docx', 'pdf']

ipcMain.handle('show-memory-file-dialog', async () => {
  return dialog.showOpenDialog(mainWindow, {
    filters: [
      { name: 'Supported Files', extensions: ALLOWED_MEMORY_EXTENSIONS },
      { name: 'Text Files', extensions: ['txt', 'md'] },
      { name: 'Office Files', extensions: ['docx'] },
      { name: 'PDF Files', extensions: ['pdf'] },
      { name: 'Data Files', extensions: ['json', 'csv'] }
    ],
    properties: ['openFile', 'multiSelections']
  })
})

ipcMain.handle('parse-memory-file', async (_event, filePath) => {
  const ext = path.extname(filePath).toLowerCase().slice(1)
  const name = path.basename(filePath)

  // Validate extension
  if (!ALLOWED_MEMORY_EXTENSIONS.includes(ext)) {
    throw new Error('不支持的文件格式: ' + ext)
  }

  // Validate path
  if (!isAllowedReadPath(filePath)) {
    throw new Error('Access to this path is not allowed')
  }

  try {
    let content = ''

    if (ext === 'txt' || ext === 'md') {
      content = await fsPromises.readFile(filePath, 'utf8')
    } else if (ext === 'json') {
      const raw = await fsPromises.readFile(filePath, 'utf8')
      const obj = JSON.parse(raw)
      content = JSON.stringify(obj, null, 2)
    } else if (ext === 'csv') {
      content = await fsPromises.readFile(filePath, 'utf8')
    } else if (ext === 'docx') {
      try {
        const mammoth = require('mammoth')
        const result = await mammoth.extractRawText({ path: filePath })
        content = result.value
      } catch (e) {
        throw new Error('解析 DOCX 失败: ' + e.message)
      }
    } else if (ext === 'pdf') {
      try {
        const pdfParse = require('pdf-parse')
        const buffer = await fsPromises.readFile(filePath)
        const data = await pdfParse(buffer)
        content = data.text
      } catch (e) {
        throw new Error('解析 PDF 失败: ' + e.message)
      }
    }

    // Truncate very large files to avoid exceeding token limits
    if (content.length > MAX_MEMORY_FILE_CHARS) {
      content = content.slice(0, MAX_MEMORY_FILE_CHARS) + `\n\n[内容已截断，显示前 ${MAX_MEMORY_FILE_CHARS} 字符]`
    }

    return {
      id: randomUUID(),
      name,
      type: ext,
      content: content.trim(),
      size: Buffer.byteLength(content, 'utf8'),
      uploadTime: new Date().toISOString()
    }
  } catch (e) {
    throw e
  }
})

ipcMain.handle('get-memory-list', () => {
  return config.memoryFiles || []
})

ipcMain.handle('save-memory-file', async (_event, fileEntry) => {
  if (!config.memoryFiles) config.memoryFiles = []
  const existing = config.memoryFiles.findIndex(f => f.id === fileEntry.id)
  if (existing >= 0) {
    config.memoryFiles[existing] = fileEntry
  } else {
    config.memoryFiles.push(fileEntry)
  }
  await saveConfig()
  return true
})

ipcMain.handle('delete-memory-file', async (_event, fileId) => {
  if (!config.memoryFiles) return true
  config.memoryFiles = config.memoryFiles.filter(f => f.id !== fileId)
  await saveConfig()
  return true
})

ipcMain.handle('update-memory-tags', async (_event, fileId, tags) => {
  if (!config.memoryFiles) return false
  const file = config.memoryFiles.find(f => f.id === fileId)
  if (file) {
    file.tags = tags
    await saveConfig()
  }
  return true
})

ipcMain.handle('open-presentation', async (_event, htmlContent) => {
  // htmlContent is already a fully self-contained presentation HTML
  // built by app.js — no injection needed here
  const tmpFile = path.join(os.tmpdir(), `ppt-pres-${Date.now()}.html`)
  await fsPromises.writeFile(tmpFile, htmlContent, 'utf8')

  const presWindow = new BrowserWindow({
    width: 1280,
    height: 720,
    fullscreen: true,
    frame: false,
    backgroundColor: '#000',
    alwaysOnTop: false,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  })

  presWindow.loadFile(tmpFile)

  // Register ESC as a safety exit even if the in-page script fails
  presWindow.webContents.on('before-input-event', (event, input) => {
    if (input.key === 'Escape') presWindow.close()
  })

  presWindow.on('closed', async () => {
    try {
      await fsPromises.unlink(tmpFile)
    } catch (_) {
      // Ignore cleanup errors
    }
  })

  return true
})

// ──── App lifecycle ────

app.whenReady().then(async () => {
  await loadConfig()
  createWindow()
  buildMenu()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})
