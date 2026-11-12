# html-FileNest

> Web-based local file manager — browse, preview, upload, and manage files from your browser. Single EXE, zero dependencies, no installation.

## Introduction

html-FileNest is a lightweight web-based file manager that runs locally on Windows. It provides a clean, modern UI (similar to VS Code's file explorer) for browsing directories, previewing files, uploading/downloading, and performing file operations — all from a web browser.

### Key Features

- 📂 **Full File Management** — Browse, create, rename, copy, move, delete, and upload files/folders
- 👁️ **Rich Preview** — Images, videos, audio, PDF, text/code (UTF-8/UTF-16/GB18030), Office (doc/docx/xls/xlsx/ppt/pptx), Design (PSD/PSB/AI), Archive browsing (ZIP/JAR/TAR/GZ/TGZ/RAR/7Z/BZ2/XZ)
- 🗂️ **Archives as Folder** — Click a ZIP/JAR/TAR/GZ/TGZ/RAR/7Z to browse its contents like a regular folder in the main file list; navigate subdirectories, preview/download files inside; archive contents are read-only (no rename/delete/copy/move) but support preview, download, and open with system handler
- 📂 **System Open** — Open any file or folder with its default system application
- 🌐 **Customizable Web UI** — Standalone `index.html` for easy personalization; dark-themed grid/list views, filename search, directory tree sidebar, sorting, and grouping
- 🔒 **Secure** — Whitelist root paths, block executables, local-origin enforcement, optional auth token
- 📁 **Full Character Support** — Perfectly handles Tibetan, Emoji, special symbols, super-long paths
- 🛡️ **Crash-Proof** — Global exception catching keeps the service alive
- ⚡ **One-Click Deploy** — Double-click to deploy, auto-start on boot, HTTPS with self-signed certificate
- 🗑️ **One-Click Uninstall** — `--uninstall` removes the service and certificates completely

## Architecture

```text
Browser (Chrome / Edge / Firefox)
    ↓ HTTP / HTTPS (127.0.0.1:20000 / 20443)
html-FileNest.exe (Node.js)
    ↓ File system operations (fs)
Local disk / Network shares
```

## Quick Start

### 1. Build

Prerequisites: Install [Bun](https://bun.sh) (recommended) or Node.js ≥ 18 + [pkg](https://github.com/vercel/pkg).

```bash
# Bun (recommended)
bun install
bun run build.ts

# Node.js + pkg (alternative)
npm install
npm install -g pkg
pkg src/html-FileNest.js --targets node18-win-x64 --output dist/html-FileNest.exe
```

### 2. Deploy

Send `html-FileNest.exe` to the target machine and double-click. It will automatically:
- Copy itself to `C:\html-FileNest\`
- Deploy `index.html` to `C:\html-FileNest\index.html` (auto-extracted from the exe; edit this file to customize the UI)
- Generate `config.json` and a self-signed HTTPS certificate
- Create a scheduled task (auto-start on login)
- Start the service and open a confirmation dialog

### 3. Open the File Manager

Navigate to **`http://127.0.0.1:20000/`** in your browser.

### 4. Verify

Open `http://127.0.0.1:20000/ping` — a JSON response with `"ok": true` means everything is running.

## API Reference

All API endpoints require requests from `127.0.0.1` / `localhost` (enforced by `requireLocalOrigin`).

### Health Check

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/ping` | Service status, version, configuration summary |

### File Operations (legacy, compatible with htmlFileAgent)

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/open?path=<abs>` | Open file with system default handler |
| GET | `/exists?path=<abs>` | Check if a file exists |
| GET | `/reveal?path=<abs>` | Reveal file in Explorer |
| GET | `/open-folder?path=<abs>` | Open containing folder in Explorer |

### File Manager API

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/roots` | List configured root directories |
| GET | `/api/drives` | List available drives (A–Z) |
| GET | `/api/ls?path=<dir>` | List directory contents |
| GET | `/api/stat?path=<abs>` | Get file/directory metadata |
| GET | `/api/raw?path=<abs>` | Stream file content (supports Range) |
| GET | `/api/raw?path=<abs>&download=1` | Download file as attachment |
| GET | `/api/read?path=<abs>` | Read text file content (max 5 MB) |
| GET | `/api/office-preview?path=<abs>` | Office/Design preview (doc/docx/xls/xlsx/ppt/pptx/psd/psb/ai) |
| GET | `/api/archive-list?path=<abs>&subdir=<dir>` | List archive contents (ZIP/JAR/TAR/GZ/TGZ/RAR/7Z, virtual directory) |
| GET | `/api/archive-extract?path=<abs>&entry=<name>` | Extract single file from archive (multi-format) |
| POST | `/api/mkdir` | Create folder `{path, name}` |
| POST | `/api/rename` | Rename `{path, newName}` |
| POST | `/api/rm` | Delete `{path, recursive?}` |
| POST | `/api/mv` | Move `{from, to}` |
| POST | `/api/cp` | Copy `{from, to}` (recursive) |
| POST | `/api/write` | Write text file `{path, content}` |
| POST | `/api/open` | Open file/folder with system handler `{path}` |
| POST | `/api/upload?path=<dir>&name=<file>` | Upload file (stream body) |

## Configuration

Edit `C:\html-FileNest\config.json`:

```json
{
  "ports": { "http": 20000, "https": 20443 },
  "mode": "whitelist",
  "sources": {
    "default": {
      "type": "local",
      "tryAll": true,
      "roots": ["*"]
    }
  },
  "allowedExtensions": [],
  "blockExecutables": true,
  "blockedExtensions": [
    ".exe", ".bat", ".cmd", ".com", ".scr", ".ps1",
    ".vbs", ".js", ".msi", ".dll", ".reg", ".lnk"
  ],
  "cacheDir": "C:\\html-FileNest\\cache",
  "cacheMaxAgeHours": 24,
  "downloadTimeoutMs": 120000,
  "allowedHosts": [],
  "security": {
    "requireLocalOrigin": true,
    "authToken": ""
  }
}
```

### Key Configuration Options

| Field | Description |
|-------|-------------|
| `ports` | HTTP/HTTPS listening ports (default: 20000/20443) |
| `sources.*.roots` | Allowed root paths. `"*"` = all drives. Use `"C:\\Users\\me\\*"` to restrict scope |
| `allowedExtensions` | Whitelist of allowed extensions. Empty array = no restriction |
| `blockExecutables` | Reject executable files (default: true) |
| `security.requireLocalOrigin` | Only accept requests from localhost (default: true) |
| `security.authToken` | Optional auth token for API access |

### Environment Variables

Config values support `${VAR}` syntax for dynamic expansion:

```json
{
  "sources": {
    "docs": {
      "type": "local",
      "roots": ["${USERPROFILE}\\Documents"]
    }
  }
}
```

## Web UI Features

- **Customizable UI** — The UI is a standalone `index.html` file. Source at `src/index.html`; deployed to `C:\html-FileNest\index.html` on first run. Edit the deployed file to personalize (hot-reloads on refresh)
- **Grid / List views** — Toggle between visual grid and detailed list
- **Directory tree sidebar** — Clickable tree showing current path hierarchy and subdirectories; "All Locations" root node for quick return
- **Filename search** — Filter files in the current directory from the toolbar
- **Sorting** — Sort by name, size, modified time, or type (ascending/descending)
- **Grouping** — Group files by category (folders, documents, images, videos, etc.)
- **File preview** — Single-click to preview images, video, audio, PDF, text/code, Office (doc/docx/xls/xlsx/ppt/pptx), Design (PSD/PSB/AI); click archives (ZIP/JAR/TAR/GZ/RAR/7Z) to browse like a folder (navigate subdirectories, preview/download files inside); drag the left edge of the preview panel to resize
- **Adaptive layout** — Filenames wrap naturally at any zoom level; list columns scale with font size
- **Preview header** — 📂 open with system handler + ⬇ download button
- **Open with system handler** — Double-click or use the 📂 button in the preview header
- **Context menu** — Right-click for Open, Preview, Reveal, Rename, Copy to…, Move to…, Delete; archive files show Preview / Download / Open with system handler
- **Drag & drop upload** — Drag files into the window to upload
- **Keyboard shortcuts** — `Backspace` (up), `Alt+←/→` (back/forward), `F5` (refresh), `Ctrl+Shift+N` (new folder), `Ctrl+U` (upload)

## Uninstall

Run in an **elevated** command prompt:
```bash
html-FileNest.exe --uninstall
```

Or manually:
1. Delete the scheduled task `html-FileNest` (taskschd.msc)
2. Kill the process and delete `C:\html-FileNest\`
3. Remove the `CN=html-FileNest Local` certificate (certlm.msc → Trusted Root Certification Authorities)

## Project Structure

```
html-FileNest/
├── src/
│   ├── html-FileNest.js    # Backend source code (Node.js/Bun)
│   └── index.html          # Standalone web UI (edit to customize)
├── dist/                   # Build output (gitignored)
│   └── html-FileNest.exe   # Compiled executable (self-contained, auto-deploys index.html)
├── build.ts                # Build script (bun run build.ts)
├── package.json            # Dependency manifest
├── bun.lock                # Dependency lock file
├── README.md
├── LICENSE
└── .gitignore
```

## FAQ

**Q: The page shows "Not Found"?**
- A: Make sure the service is running. Check `http://127.0.0.1:20000/ping`.

**Q: HTTPS certificate warning?**
- A: Self-signed certificates trigger browser warnings. This is expected for local use.

**Q: How do I restrict which directories are accessible?**
- A: Edit `sources.default.roots` in `config.json`, e.g. `["D:\\Docs", "\\\\server\\share"]`.

**Q: How do I restrict file types?**
- A: Set `allowedExtensions`, e.g. `[".pdf", ".docx", ".xlsx", ".txt"]`.

**Q: API returns "origin not allowed"?**
- A: The request's `Origin` header is not localhost. For same-origin requests from the built-in UI, this should not occur.

## Changelog

**Unreleased**
- License compliance: full third-party license texts (Bun, JSZip, mammoth, word-extractor, SheetJS, 7-Zip notice) are now embedded inside the executable. Run `html-FileNest.exe --licenses` to display them at any time.
- Docs: corrected the SheetJS license note — v0.20.3 (SheetJS CDN) is Apache-2.0, same license as before; the previous "newer versions use a different license" statement was outdated.

**v1.4.2.1 (Latest)**
- Security: fixed command injection in `/api/open`. The endpoint used `cmd.exe /c start`, where `cmd.exe` interprets shell metacharacters (`&`, `|`, …) — a crafted filename could execute arbitrary commands. It now launches files via `explorer.exe` (argv passing, no shell interpretation) with a `rundll32 url.dll,FileProtocolHandler` fallback, matching the `/open` file-op implementation.
- `/api/open` now rejects paths/entries containing Windows-invalid filename characters (quotes, `<>|*?`, control chars) as an extra hardening layer. Legal metacharacters (`&`, `%`, …) remain allowed since they are inert in argv passing.

**v1.4.1.2**
- Security: `blockedExtensions` is now enforced on `/api/raw`, `/api/read`, `/api/archive-extract` and `/api/open`. Previously only the `/open`/`/reveal` file ops checked it, so blocked executables could still be downloaded, read or opened directly through the API.
- `checkExtension` now ignores trailing dots/spaces before matching (prevents `evil.exe.` / `evil.exe ` bypasses, since Windows silently strips them).

**v1.4.1.1**
- Security: upgraded `xlsx` to SheetJS 0.20.3, fixing GHSA-4r6h-8v6p-xvw6 (prototype pollution, CVSS 7.8) and GHSA-5pgg-2g8v-p4x9 (ReDoS, CVE-2024-22363, CVSS 7.5).
- The `xlsx` dependency is intentionally pinned to the SheetJS CDN tarball (`https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`). npm's `xlsx` package is frozen at 0.18.5 and never receives these fixes — do NOT switch back to an npm version range, and do NOT merge Dependabot "upgrade xlsx" PRs (they reintroduce the vulnerabilities).

**v1.4.0**
- Multi-format archive support: added RAR, 7Z (via 7-Zip CLI), TAR (native parser), GZ/TGZ (zlib + TAR), JAR (JSZip) — all open as navigable folders like ZIP.
- Unified archive interface: `/api/archive-list` and `/api/archive-extract` now dispatch by format (JSZip / TAR parser / 7z CLI).
- Archive entries support preview, download, and system-handler open across all formats.

**v1.3.0**
- Added legacy Office format preview: `.doc` (text extraction via word-extractor), `.xls` (via SheetJS), `.ppt` (binary text scan).
- Fixed non-ASCII (Chinese/CJK) filename preview failure: `Content-Disposition` header now uses RFC 5987 dual encoding to avoid `ERR_INVALID_CHAR`.
- Extension normalization: multi-dot filenames (`.tar.gz` → `gz`) via `lastIndexOf('.')`.
- Image format support expanded to 17 formats (added `jpe`, `jfif`, `heic`, `heif`, `apng`).

**v1.2.0**
  - ZIP archives now open as folders in the main file list (click to enter, navigate subdirectories, preview/download inside) instead of the preview pane.
  - ZIP contents are read-only (rename/delete/copy/move disabled) but support preview, download, and open with system handler (extracted to temp).
  - Breadcrumb shows ZIP virtual path navigation.

  **v1.1.0**
- Fixed AI file preview: PDF-compatible AI renders in iframe via Blob URL; PostScript AI shows system-handler prompt.
- Fixed browser auto-downloading raw files on preview: added `Content-Disposition: inline` for non-download responses.
- Removed 20MB file size limit for PSD/AI/Office preview.
- PSD composite image extraction: added ZIP compression (2/3) and 16-bit depth support; fixed header offset parsing.
- Added ZIP archive browsing (navigate subfolders, preview/download files inside) and `/api/open` system-handler endpoint.
- Pinned dependency versions (xlsx 0.18.5, jszip 3.10.2, mammoth 1.13.0) for license compliance.

**v1.0.0**
- Initial release.
- Compatible with htmlFileAgent's `/open`, `/exists`, `/reveal`, `/open-folder` endpoints.

## Acknowledgments

### Own Projects

- **[htmlFileAgent](https://github.com/FatFatYoung/htmlFileAgent)** — html-FileNest inherits and extends htmlFileAgent's file-opening architecture, including the `/open`, `/exists`, `/reveal`, `/open-folder` API design, multi-source `roots` path resolution, wildcard matching, crash-proof armor pattern, and deployment model (single EXE, double-click deploy, `--uninstall`).

### Third-Party Open Source

| Project | Role | License |
|---------|------|---------|
| [Node.js](https://nodejs.org) | Runtime environment (fs, http, https, crypto, child_process, etc.) | MIT |
| [Bun](https://bun.sh) | Recommended build toolchain & single-EXE compiler | MIT (bundles JavaScriptCore, LGPL-2) |
| [pkg](https://github.com/vercel/pkg) | Alternative Node.js-to-EXE compiler | MIT |
| [mammoth](https://www.npmjs.com/package/mammoth) | DOCX → HTML conversion for document preview | BSD-2-Clause |
| [word-extractor](https://www.npmjs.com/package/word-extractor) | Legacy .doc (OLE2) text extraction | MIT |
| [SheetJS (xlsx)](https://sheetjs.com) | XLSX/XLS spreadsheet parsing for table preview | Apache-2.0 |
| [JSZip](https://stuk.github.io/jszip/) | ZIP/JAR archive browsing/extract + PPTX slide text parsing | MIT/GPL-3.0 |
| [7-Zip](https://www.7-zip.org) | RAR/7Z/BZ2/XZ archive listing & extraction (CLI, runtime dependency) | LGPL-2.1+ |

> **Note**: Office document "open with system handler" uses the OS file association (`explorer.exe`/`cmd /c start`) — no third-party Office software is bundled or required. SheetJS is pinned to v0.20.3 from the official SheetJS CDN (Apache-2.0). Full third-party license texts are embedded in the executable and can be displayed with `html-FileNest.exe --licenses`.

### Standards & Specifications

- **RFC 7233** — HTTP Range Requests (used for `/api/raw` video/audio streaming)
- **RFC 6838** — Media Type specifications (MIME mapping in `/api/raw`)
- **BOM / Unicode** — UTF-8 (EF BB BF), UTF-16 LE (FF FE), UTF-16 BE (FE FF) encoding auto-detection
- **GB18030 / GBK** — Chinese text encoding fallback when UTF-8 decoding produces replacement characters

## License

[MIT](LICENSE) © 2026 [FatFatYoung](https://github.com/FatFatYoung)

## Author

**FatFatYoung** — [github.com/FatFatYoung](https://github.com/FatFatYoung)
