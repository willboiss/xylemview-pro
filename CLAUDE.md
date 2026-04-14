# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

- `npm start` — Run the app (Electron)
- `npm run dist:win` — Build Windows installer (NSIS, per-user)
- `npm run dist:mac` — Build macOS DMG

## Architecture

XylemView Pro is a frameless Electron desktop app with acrylic/glass UI for Xylem engineering staff. It's a file/order/drawing management tool that replaces manual Windows Explorer navigation.

**Three files make up the entire app:**

- **main.js** (~4,500 lines) — Main process. All business logic, IPC handlers (~110), file system operations, SOAP client, PCOMM automation, Claude AI streaming. All network I/O is async (fs.promises) to prevent UI freezing on VPN.
- **index.html** (~8,000 lines) — Single-file renderer. All CSS, JS, and HTML inline. No bundler, no frameworks, no frontend npm deps. State is managed through a global `S` object that drives all UI renders via direct DOM manipulation.
- **preload.js** (~165 lines) — Context bridge exposing `window.api.*` to the renderer. This is the API contract between main and renderer.

Also: `contingency-worker.js` (worker thread for MDB database parsing).

### IPC Pattern

All cross-process communication flows through preload.js. Adding a new feature means:
1. Add `ipcMain.handle('channel-name', ...)` in main.js
2. Expose it in preload.js as `ipcRenderer.invoke('channel-name', ...)`
3. Call it in index.html as `await window.api.channelName(...)`

### State Management

The renderer's global `S` object holds all UI state: current mode, files, search results, config, OTP digits, etc. When switching between tabs (Orders/Drawings/Marketing/Chat), state is saved/restored via `saveModeState()`/`restoreModeState()` so search results persist across tab switches.

### Themes

CSS variables power 5 themes: dark, light, tron, disco, classic. Applied via `data-theme` attribute on `<body>`. Theme-specific CSS uses `[data-theme="X"]` selectors.

## Critical Rules

### Safety Policy — NEVER delete or overwrite files
- Link creation uses `fs.writeFile` with `wx` flag (exclusive create)
- File paste uses `COPYFILE_EXCL`
- No rm, no unlink, no overwrite of user files. This is a hard constraint.

### Legacy Format Caution
Xylem runs extremely old infrastructure (VB6, Jet 3.x, DAO 3.5, AS/400). When working with MDB/MDE files, **always** preserve Jet 3.x format (`Jet OLEDB:Engine Type=4`). Never silently upgrade database formats — it breaks downstream VB6 apps.

### API Keys
`claude-key.txt` is gitignored and must never be committed. The app reads it from a local file or network share at runtime.

## Network Paths

The app depends on corporate network drives (or UNC fallbacks when drive letters aren't mapped):
- Orders: `L:\Group\orders\` / `\\01ckfp02-1\vol1\Group\orders\`
- Drawings: `L:\Drawings\ACADDWGS\`
- Shared app data: `E:\XylemView\XylemView Pro\` (chat.json, nicknames, installer, ODA converter)
- Contingency DBs: multiple MDB files on E: drive, watched via `fs.watch()`

## Drawing Number Convention

Full drawing number: `######X#####` (12 chars), e.g., `123456A78901`. Revision suffix: `rXX`. Link files use square brackets: `[123456A78901r0A.dwg]` with content pointing to the real file path.

## Deploy Process

1. Bump version in three places: `CURRENT_VERSION` in main.js, `version` in package.json, and the `CHANGELOG` array in index.html
2. Run `npm run dist:win`
3. Copy installer to `E:\XylemView\XylemView Pro\`
4. Users see update banner automatically (app checks installer ProductVersion on the network share)
