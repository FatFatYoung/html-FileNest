# html-FileNest

> Web-based local file manager — browse, preview, upload, and manage files from your browser. Single EXE, zero dependencies, no installation.

## Introduction

html-FileNest is a lightweight web-based file manager that runs locally on Windows. It provides a clean, modern UI (similar to VS Code's file explorer) for browsing directories, previewing files, uploading/downloading, and performing file operations — all from a web browser.

### Key Features

- 📂 **Full File Management** — Browse, create, rename, copy, move, delete, and upload files/folders
- 👁️ **Rich Preview** — Images, videos, audio, PDF, text/code files (UTF-8/UTF-16 auto-detect), and Office documents (via WPS)
- 🌐 **Customizable Web UI** — Standalone `index.html` for easy personalization; dark-themed grid/list views, breadcrumb, search, sorting, and grouping
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
bun build --compile src/html-FileNest.js --outfile dist/html-FileNest.exe

# Node.js + pkg (alternative)
npm install -g pkg
pkg src/html-FileNest.js --targets node18-win-x64 --output dist/html-FileNest.exe
```

### 2. Deploy

Send `html-FileNest.exe` to the target machine and double-click. It will automatically:
- Copy itself to `C:\html-FileNest\`
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
| POST | `/api/mkdir` | Create folder `{path, name}` |
| POST | `/api/rename` | Rename `{path, newName}` |
| POST | `/api/rm` | Delete `{path, recursive?}` |
| POST | `/api/mv` | Move `{from, to}` |
| POST | `/api/cp` | Copy `{from, to}` (recursive) |
| POST | `/api/write` | Write text file `{path, content}` |
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

- **Customizable UI** — The UI is a standalone `src/index.html` file; edit it directly to personalize the look and feel (hot-reloads on refresh)
- **Grid / List views** — Toggle between visual grid and detailed list
- **Breadcrumb navigation** — Click any path segment to jump
- **Search filter** — Filter files in the current directory
- **Sorting** — Sort by name, size, modified time, or type (ascending/descending)
- **Grouping** — Group files by category (folders, documents, images, videos, etc.)
- **File preview** — Single-click to preview images, video, audio, PDF, text/code, Office docs
- **Open with system handler** — Double-click a file to open it with the system default application
- **Context menu** — Right-click for Open, Preview, Reveal, Rename, Copy to…, Move to…, Delete
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
│   ├── html-FileNest.exe   # Compiled executable
│   └── index.html          # Copied from src/ during build
├── build.ts                # Build script (bun build.ts)
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

**v1.0.0 (Latest)**
- Initial release: Web-based local file manager with full file operations, rich preview, and secure access control.
- Standalone `index.html` UI — easily customizable without touching backend code (hot-reloads on refresh).
- Smart text preview with UTF-8/UTF-16 auto-detection (BOM-aware, no more false "binary file" errors).
- Double-click opens files with the system default handler; preview via single-click or context menu.
- File copy (`/api/cp`, recursive) and move operations with dialog-based target selection.
- Client-side sorting (name/size/time/type, asc/desc) and grouping by file category.
- Compatible with htmlFileAgent's `/open`, `/exists`, `/reveal`, `/open-folder` endpoints.

## Acknowledgments

### Own Projects

- **[htmlFileAgent](https://github.com/FatFatYoung/htmlFileAgent)** — html-FileNest inherits and extends htmlFileAgent's file-opening architecture, including the `/open`, `/exists`, `/reveal`, `/open-folder` API design, multi-source `roots` path resolution, wildcard matching, crash-proof armor pattern, and deployment model (single EXE, double-click deploy, `--uninstall`).

### Third-Party Open Source

| Project | Role | License |
|---------|------|---------|
| [Node.js](https://nodejs.org) | Runtime environment (fs, http, https, crypto, child_process, etc.) | MIT |
| [Bun](https://bun.sh) | Recommended build toolchain & single-EXE compiler | MIT / Apache-2.0 |
| [pkg](https://github.com/vercel/pkg) | Alternative Node.js-to-EXE compiler | MIT |
| [WPS Office](https://www.wps.cn) | Office document preview engine (via browser IFrame) | Proprietary (free tier) |

### Standards & Specifications

- **RFC 7233** — HTTP Range Requests (used for `/api/raw` video/audio streaming)
- **RFC 6838** — Media Type specifications (MIME mapping in `/api/raw`)
- **BOM / Unicode** — UTF-8 (EF BB BF), UTF-16 LE (FF FE), UTF-16 BE (FE FF) encoding auto-detection

## License

[MIT](LICENSE) © 2026 [FatFatYoung](https://github.com/FatFatYoung)

## Author

**FatFatYoung** — [github.com/FatFatYoung](https://github.com/FatFatYoung)
