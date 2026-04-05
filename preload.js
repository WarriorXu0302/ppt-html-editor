const { contextBridge, ipcRenderer } = require('electron')

// Track registered listeners to prevent duplicates
const registeredListeners = new Map()

contextBridge.exposeInMainWorld('electronAPI', {
  platform: process.platform,

  // File I/O
  readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
  writeFile: (filePath, content) => ipcRenderer.invoke('write-file', filePath, content),

  // Dialogs
  showOpenDialog: () => ipcRenderer.invoke('show-open-dialog'),
  showSaveDialog: (defaultPath) => ipcRenderer.invoke('show-save-dialog', defaultPath),
  showMessageBox: (options) => ipcRenderer.invoke('show-message-box', options),

  // Config
  getConfig: () => ipcRenderer.invoke('get-config'),
  setConfig: (updates) => ipcRenderer.invoke('set-config', updates),

  // Paths
  getPath: (name) => ipcRenderer.invoke('get-path', name),

  // Window
  setTitle: (title) => ipcRenderer.send('set-title', title),
  setDocumentEdited: (edited) => ipcRenderer.send('set-document-edited', edited),
  setDirtyFlag: (dirty) => ipcRenderer.send('set-dirty-flag', dirty),

  // Save completion notification
  notifySaveComplete: () => ipcRenderer.send('save-complete'),

  // Presentation window
  openPresentation: (htmlContent) =>
    ipcRenderer.invoke('open-presentation', htmlContent),

  // PPTX import
  openPptxFile: () => ipcRenderer.invoke('open-pptx-file'),

  // Memory files
  showMemoryFileDialog: () => ipcRenderer.invoke('show-memory-file-dialog'),
  parseMemoryFile: (filePath) => ipcRenderer.invoke('parse-memory-file', filePath),
  getMemoryList: () => ipcRenderer.invoke('get-memory-list'),
  saveMemoryFile: (fileEntry) => ipcRenderer.invoke('save-memory-file', fileEntry),
  deleteMemoryFile: (fileId) => ipcRenderer.invoke('delete-memory-file', fileId),
  updateMemoryTags: (fileId, tags) => ipcRenderer.invoke('update-memory-tags', fileId, tags),

  // Menu event listeners with cleanup support
  onMenuEvent: (channel, callback) => {
    const validChannels = [
      'menu-open', 'menu-save', 'menu-save-as', 'menu-undo',
      'menu-redo', 'menu-open-file', 'menu-new'
    ]
    if (!validChannels.includes(channel)) {
      console.warn(`[preload] Invalid menu channel: ${channel}`)
      return () => {}
    }

    // Remove existing listener for this channel to prevent duplicates
    if (registeredListeners.has(channel)) {
      const oldListener = registeredListeners.get(channel)
      ipcRenderer.removeListener(channel, oldListener)
    }

    // Create wrapped listener
    const wrappedCallback = (event, ...args) => callback(...args)
    registeredListeners.set(channel, wrappedCallback)
    ipcRenderer.on(channel, wrappedCallback)

    // Return cleanup function
    return () => {
      ipcRenderer.removeListener(channel, wrappedCallback)
      registeredListeners.delete(channel)
    }
  },

  removeMenuListener: (channel) => {
    if (!registeredListeners.has(channel)) {
      console.warn(`[preload] No listener registered for: ${channel}`)
      return
    }
    const listener = registeredListeners.get(channel)
    ipcRenderer.removeListener(channel, listener)
    registeredListeners.delete(channel)
  }
})
