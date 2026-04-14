const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  // Config
  getConfig:       ()           => ipcRenderer.invoke('get-config'),
  saveConfig:      (cfg)        => ipcRenderer.invoke('save-config', cfg),
  getSystemDark:   ()           => ipcRenderer.invoke('get-system-dark'),
  getVersion:      ()           => ipcRenderer.invoke('get-version'),
  quitApp:         ()           => ipcRenderer.invoke('quit-app'),
  logError:        (msg)        => ipcRenderer.invoke('log-renderer-error', msg),
  getBirthdays:    ()           => ipcRenderer.invoke('get-birthdays'),
  getPlatform:     ()           => ipcRenderer.invoke('get-platform'),

  // Network
  checkNetwork:    ()            => ipcRenderer.invoke('check-network'),
  getPathDefaults: ()            => ipcRenderer.invoke('get-path-defaults'),

  // Search
  searchOrder:     (order, line) => ipcRenderer.invoke('search-order', order, line),
  deepSearchOrder: (order, line) => ipcRenderer.invoke('deep-search-order', order, line),
  searchDrawing:   (input)       => ipcRenderer.invoke('search-drawing', input),

  // File operations
  openFile:        (fp)         => ipcRenderer.invoke('open-file', fp),
  openInViewer:    (fp)         => ipcRenderer.invoke('open-in-viewer', fp),
  openInAutocad:   (fp)         => ipcRenderer.invoke('open-in-autocad', fp),
  detectAutocad:   ()           => ipcRenderer.invoke('detect-autocad'),
  checkAccore:     ()           => ipcRenderer.invoke('check-accore'),
  checkDxfConverter: ()         => ipcRenderer.invoke('check-dxf-converter'),
  openLink:        (fp)         => ipcRenderer.invoke('open-link', fp),
  readLink:        (fp)         => ipcRenderer.invoke('read-link', fp),
  openFolder:      (fp)         => ipcRenderer.invoke('open-folder', fp),
  showInFolder:    (fp)         => ipcRenderer.invoke('show-in-folder', fp),

  // Nicknames (shared username → display name map)
  getNicknames:       ()         => ipcRenderer.invoke('get-nicknames'),

  // Contingency
  preloadContingency: (force)    => ipcRenderer.invoke('preload-contingency', force),
  getContingency:    (ol)       => ipcRenderer.invoke('get-contingency', ol),
  scanChecklists:    (order)    => ipcRenderer.invoke('scan-checklists', order),
  scanNewOrders:     ()         => ipcRenderer.invoke('scan-new-orders'),
  openContingencyStandalone: () => ipcRenderer.invoke('open-contingency-standalone'),
  openContingencyOrder: (id)    => ipcRenderer.invoke('open-contingency-order', id),
  printContingency:    (html)           => ipcRenderer.invoke('print-contingency', html),
  saveContingencyPdf:  (orderLine, html) => ipcRenderer.invoke('save-contingency-pdf', orderLine, html),
  onContingencyLoading: (cb)             => ipcRenderer.on('contingency-loading', (_, loading) => cb(loading)),
  onPdfPaperDetected:  (cb)              => ipcRenderer.on('pdf-paper-detected', (_, letter, orient) => cb(letter, orient)),

  // Link to order
  createLink: (src, folder, create) => ipcRenderer.invoke('create-link', src, folder, create),

  // Clipboard
  copyFileToClipboard: (fp)     => ipcRenderer.invoke('copy-file-to-clipboard', fp),
  pasteFiles:     (folder)      => ipcRenderer.invoke('paste-files', folder),
  copyText:       (text)        => ipcRenderer.invoke('copy-text', text),

  // Misc
  sendEmail:      (addr, subj)  => ipcRenderer.invoke('send-email', addr, subj),
  sendEmailWithBody: (addr, subj, body) => ipcRenderer.invoke('send-email-with-body', addr, subj, body),
  sendEmailHtml: (addr, subj, html) => ipcRenderer.invoke('send-email-html', addr, subj, html),
  openExternal:   (url)         => ipcRenderer.invoke('open-external', url),
  wingetInstall:  (id)          => ipcRenderer.invoke('winget-install', id),
  restartApp:     ()            => ipcRenderer.invoke('restart-app'),
  openMarketing:  (o)           => ipcRenderer.invoke('open-marketing', o),
  findChecklist:  (o)           => ipcRenderer.invoke('find-checklist', o),
  getSiblingLines:(o)           => ipcRenderer.invoke('get-sibling-lines', o),

  // PDF preview
  previewPdf:      (fp, title)   => ipcRenderer.invoke('preview-pdf', fp, title),
  checkLinksExist: (order, items) => ipcRenderer.invoke('check-links-exist', order, items),

  // Ask Claude
  askClaude:       (order)       => ipcRenderer.invoke('ask-claude', order),
  setClaudeKey:    (key)         => ipcRenderer.invoke('set-claude-key', key),
  onClaudeChunk:   (cb)          => ipcRenderer.on('claude-chunk', (_, text) => cb(text)),

  // Recent orders
  getRecentOrders: ()           => ipcRenderer.invoke('get-recent-orders'),
  addRecentOrder: (o, l)        => ipcRenderer.invoke('add-recent-order', o, l),
  getRecentMktOrders: ()        => ipcRenderer.invoke('get-recent-mkt-orders'),
  addRecentMktOrder: (o)        => ipcRenderer.invoke('add-recent-mkt-order', o),

  // Recent drawings
  getRecentDrawings: ()         => ipcRenderer.invoke('get-recent-drawings'),
  addRecentDrawing: (q)         => ipcRenderer.invoke('add-recent-drawing', q),

  // Rev-checking
  checkRevisions: (files)       => ipcRenderer.invoke('check-revisions', files),

  // PDF preview
  findPdfPreview: (fp)          => ipcRenderer.invoke('find-pdf-preview', fp),

  // Disco song
  pickDiscoSong:  ()            => ipcRenderer.invoke('pick-disco-song'),
  getDiscoSong:   ()            => ipcRenderer.invoke('get-disco-song'),

  // User info
  getUserInfo:    ()            => ipcRenderer.invoke('get-user-info'),

  // Rev update
  updateRevLinks: (u)           => ipcRenderer.invoke('update-rev-links', u),

  // Auto-update
  checkUpdate:    ()            => ipcRenderer.invoke('check-update'),
  openInstaller:  ()            => ipcRenderer.invoke('open-installer'),

  // Konami faces + cursor
  getRecentOpenedDrawings: ()    => ipcRenderer.invoke('get-recent-opened-drawings'),
  addRecentOpenedDrawing: (n, q) => ipcRenderer.invoke('add-recent-opened-drawing', n, q),
  getKonamiFaces: ()            => ipcRenderer.invoke('get-konami-faces'),
  getKonamiCursor: ()           => ipcRenderer.invoke('get-konami-cursor'),

  // BOM Extract
  extractBom:     (fp)          => ipcRenderer.invoke('extract-bom-from-dwg', fp),
  saveBom:        (data)        => ipcRenderer.invoke('save-bom', data),

  // Resubmit BOM/DRS
  resubmitFile:   (fp)          => ipcRenderer.invoke('resubmit-file', fp),

  // DWG → DXF / PDF
  convertDwgDxf:  (fp, outDir)  => ipcRenderer.invoke('convert-dwg-dxf', fp, outDir),
  convertDwgPdf:  (fp, outDir, opts)  => ipcRenderer.invoke('convert-dwg-pdf', fp, outDir, opts),
  getDownloadsPath: ()          => ipcRenderer.invoke('get-downloads-path'),
  scanTubesheet:  (fp)          => ipcRenderer.invoke('scan-tubesheet', fp),

  // Routing
  parseDrsGphe:      (fp)     => ipcRenderer.invoke('parse-drs-gphe', fp),
  parseDrsBrazed:    (fp)     => ipcRenderer.invoke('parse-drs-brazed', fp),
  scanFolderUtube:   (folder) => ipcRenderer.invoke('scan-folder-utube', folder),
  convertDwgDxfText: (fp)     => ipcRenderer.invoke('convert-dwg-dxf-text', fp),
  cancelConvert:    ()        => ipcRenderer.invoke('cancel-convert'),

  // Chat (Firebase real-time)
  chatRead:       ()            => ipcRenderer.invoke('chat-read'),
  chatSend:       (text)        => ipcRenderer.invoke('chat-send', text),
  chatReact:      (msgId, emoji) => ipcRenderer.invoke('chat-react', msgId, emoji),
  chatEdit:       (msgId, text)  => ipcRenderer.invoke('chat-edit', msgId, text),
  chatUnsend:     (msgId)        => ipcRenderer.invoke('chat-unsend', msgId),
  chatSetPfp:     (dataUrl)      => ipcRenderer.invoke('chat-set-pfp', dataUrl),
  chatGetPfps:    ()             => ipcRenderer.invoke('chat-get-pfps'),
  onChatUpdate:   (cb)          => ipcRenderer.on('chat-update', (_, msgs) => cb(msgs)),
  notifyChat:     (title, body) => ipcRenderer.invoke('notify-chat', title, body),
  isWindowFocused: ()           => ipcRenderer.invoke('is-window-focused'),
  commentsRead:   (ol)          => ipcRenderer.invoke('comments-read', ol),
  commentsSend:   (ol, text)    => ipcRenderer.invoke('comments-send', ol, text),

  // Route-O-Matic (PCOMM)
  routeOMatic:    (ops, session)  => ipcRenderer.invoke('route-o-matic', ops, session),
  readPcommItem:  (session)       => ipcRenderer.invoke('read-pcomm-item', session),
  routeOMaticAcs: (ops)           => ipcRenderer.invoke('route-o-matic-acs', ops),

  // Diagnostics
  collectDiagnostics: ()        => ipcRenderer.invoke('collect-diagnostics'),
  logDiag:       (cat, msg)     => ipcRenderer.invoke('log-diag', cat, msg),

  // BPCS web service (read-only test methods)
  bpcsGetItemNumber:      (ct, order, line) => ipcRenderer.invoke('bpcs-get-item-number', ct, order, line),
  bpcsCheckPnExists:        (ct, pn)        => ipcRenderer.invoke('bpcs-check-pn-exists', ct, pn),
  bpcsCheckExistingBom:     (ct, pn)        => ipcRenderer.invoke('bpcs-check-existing-bom', ct, pn),
  bpcsCheckExistingRouters: (ct, pn)        => ipcRenderer.invoke('bpcs-check-existing-routers', ct, pn),
  bpcsGetRouters:         (pn)              => ipcRenderer.invoke('bpcs-get-routers', pn),
  bpcsOrderEom:           (order)         => ipcRenderer.invoke('bpcs-order-eom', order),
  bpcsBacklogCheckin:     ()              => ipcRenderer.invoke('bpcs-backlog-checkin'),
  bpcsDrawingHistory:     (dwg)           => ipcRenderer.invoke('bpcs-drawing-history', dwg),
  bpcsCheckExistingFrtLine: (ct, pn, op) => ipcRenderer.invoke('bpcs-check-existing-frt-line', ct, pn, op),
  bpcsInsertFrtTest:      (pn, op, wc, desc) => ipcRenderer.invoke('bpcs-insert-frt-test', pn, op, wc, desc),
  bpcsUploadRouting:      (ct, pn, ops)       => ipcRenderer.invoke('bpcs-upload-routing', ct, pn, ops),
  bpcsMaxOrder:           ()                 => ipcRenderer.invoke('bpcs-max-order'),

  // EOM direct (jt400 Java helper)
  eomCheckin:     (contract, status, pename) => ipcRenderer.invoke('eom-checkin', contract, status, pename),
  eomSetDates:    (contract, status, pe, df, bm, me) => ipcRenderer.invoke('eom-set-dates', contract, status, pe, df, bm, me),

  // Saturn backlog API
  saturnLogin:    (user, pass)        => ipcRenderer.invoke('saturn-login', user, pass),
  saturnStatus:   ()                  => ipcRenderer.invoke('saturn-status'),
  saturnBacklog:  (filter, size, pg)  => ipcRenderer.invoke('saturn-backlog', filter, size, pg),
  saturnOrder:    (eid, contract)     => ipcRenderer.invoke('saturn-order', eid, contract),
  saturnUsers:    ()                  => ipcRenderer.invoke('saturn-users'),
  saturnSchedule: (action, payload)   => ipcRenderer.invoke('saturn-schedule', action, payload),

  // Window controls
  winMinimize:    ()            => ipcRenderer.invoke('win-minimize'),
  winMaximize:    ()            => ipcRenderer.invoke('win-maximize'),
  winClose:       ()            => ipcRenderer.invoke('win-close'),
  winIsMaximized: ()            => ipcRenderer.invoke('win-is-maximized'),
  setGlassMode:  (on)           => ipcRenderer.invoke('set-glass-mode', on),
  onSystemThemeChanged: (cb)    => ipcRenderer.on('system-theme-changed', cb),
  onMaximizedChanged:   (cb)    => ipcRenderer.on('maximized-changed', (_, maximized) => cb(maximized)),
});
