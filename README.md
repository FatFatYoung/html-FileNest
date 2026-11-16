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
- 🔄 **Service Control** — `--restart` / `--stop` commands, plus a restart button at the bottom of the left sidebar with icon and label (for config/ certificate changes and upgrades)
- 🤖 **AI-Ready API** — A file interface with a damage ceiling for machine callers: caller allowlist, scoped roots, path blacklist, recycle-by-default (see "AI-Ready File Interface" below)

## AI-Ready File Interface

When an AI agent manipulates files, the real question is not speed — it is the blast radius. Direct file access hands the agent the whole machine and relies on prompt-level discipline; html-FileNest puts **code-enforced boundaries** between the agent and your disk. Its REST API was built for machine callers from day one — the legacy `/ping`, `/open`, `/exists`, `/reveal`, `/open-folder` endpoints are kept compatible with [htmlFileAgent](https://github.com/FatFatYoung/htmlFileAgent), and the full `/api/*` surface (list, read, write, rename, move, upload, recycle, search, ...) is a stable contract for automation.

### What the API layer gives you that direct file access cannot

Direct file access has no notion of *who is calling*, *where they may reach*, or *how to take it back*. The HTTP layer adds all three — enforced by code, not by good intentions:

- 🛂 **Caller identity** — Every request is gated by the webpage allowlist (`security.allowedOrigins`): a forged `Host`, a cross-site `Origin`, or `Origin: null` gets 403 before any file is touched. The app's own UI is always allowed; every other caller must be registered explicitly. Direct file access has no equivalent gate at all.
- 🎯 **Scope** — `roots` bounds where callers can reach, and `deniedPaths` hides and locks sensitive trees entirely — the install directory holding `config.json`, certificates and cache is always denied and cannot be un-denied (the only exception is the app-managed recycle bin, which stays readable for the recycle UI's preview and download — writes remain blocked), and junction/symlink redirects into denied trees are resolved to real paths and blocked.
- ♻️ **Reversibility** — `/api/rm` moves items to the app-level recycle bin by default and only truly deletes with an explicit `purge: true`; blocked executable extensions are enforced on every write path (upload, rename, move, copy, save, new-file creation). Even a fully tricked agent has a damage ceiling.

### Where it fits

| Caller | Recommended path |
| --- | --- |
| Trusted local agent with a human in the loop | Direct file access is fine and faster |
| Web pages, remote clients, third-party agents | Through the API — the allowlist gate exists exactly for this |
| Unattended batch jobs | Through the API — recycle-by-default keeps every change recoverable |

The point is not "an AI platform" — it is the **boundary**: a bounded, reversible file surface for callers you do not fully trust, while full-power direct access stays reserved for callers you do.

> **Integration discipline for agents**: when the API is unavailable, stop and report — never fall back to raw shell access. A fallback replaces the code-enforced boundaries above with prompt-level discipline and voids every guarantee at once.

### Development direction

- ✂️ **Precise edit (`edit`/`patch`)** — change a few lines without rewriting the whole file; whole-file overwrite amplifies the risk surface (a concurrent change between read and write gets clobbered), so this gap comes first
- 📒 **Structured audit log (JSONL)** — every API call recorded with caller, action, result and timing, so "who touched what" is always answerable
- 🔁 **Idempotency keys for mutating endpoints** — a retried or timed-out delete can never delete twice
- 🔌 **MCP adapter layer** — translate standard MCP file tools onto the existing REST API so Claude/Cursor-class clients connect with zero adaptation, while the human UI and the htmlFileAgent compatibility layer stay untouched

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

All API endpoints are protected by a webpage allowlist: the app's own UI origin - localhost/127.0.0.1 on the ports configured in `config.json` (hot-reloaded, so changing `ports` takes effect without restart; defaults `http://127.0.0.1:20000`, `https://127.0.0.1:20443`) is always allowed and is not subject to the list; same-origin calls whose `Origin` host matches the request `Host` are always allowed as well. Other sites must be registered in `security.allowedOrigins`, either as a full origin (`http://192.168.1.10` - exact scheme, host and port) or as a bare domain (`192.168.1.10` - exact hostname, any scheme or port). Matching is exact: no wildcards and no subdomain inheritance - `b.a.com` needs its own entry even when `a.com` is listed. An empty list (default) rejects every other site; `Origin: null` is always rejected. Requests with no `Origin` (curl, local scripts) are allowed.

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
| GET | `/api/desktop` | List the user's Desktop folder (path + entries) |
| GET | `/api/ls?path=<dir>&sort=<key>&asc=<0-1>&fast=<0-1>` | List directory contents; sort=key is name/size/mtime/type, asc=0 or 1 (source-side sort), fast=1 skips per-file stat |
| GET | `/api/stat?path=<abs>` | Get file/directory metadata |
| POST | `/api/stat-batch` | Batched stat lookup `{paths: []}` (max 2000 per request, read-only) |
| GET | `/api/raw?path=<abs>` | Stream file content (supports Range) |
| GET | `/api/raw?path=<abs>&download=1` | Download file as attachment |
| GET | `/api/read?path=<abs>` | Read text file content (max 5 MB) |
| GET | `/api/office-preview?path=<abs>` | Office/Design preview (doc/docx/xls/xlsx/ppt/pptx/psd/psb/ai) |
| GET | `/api/archive-list?path=<abs>&subdir=<dir>` | List archive contents (ZIP/JAR/TAR/GZ/TGZ/RAR/7Z, virtual directory) |
| GET | `/api/archive-extract?path=<abs>&entry=<name>` | Extract single file from archive (multi-format) |
| POST | `/api/mkdir` | Create folder `{path, name}` |
| POST | `/api/mkfile` | Create an empty file `{path, name}` (fails if the name exists) |
| GET | `/api/newfile-types` | File types registered in the Windows "New" context menu (registry `ShellNew`) |
| GET | `/api/favorites` | List favorite folders/files |
| POST | `/api/favorites` | Add a favorite `{path, type}` |
| POST | `/api/unfavorite` | Remove a favorite `{path}` |
| POST | `/api/rename` | Rename `{path, newName}` |
| POST | `/api/rm` | Move to recycle bin `{path, recursive?}` (truly deletes only with `purge: true`) |
| GET | `/api/recycle` | List recycle-bin items (original path, deletion time) |
| POST | `/api/recycle-restore` | Restore a recycled item `{id}` (name collisions get a " (restored)" suffix) |
| POST | `/api/recycle-purge` | Permanently delete one recycled item `{id}` |
| POST | `/api/recycle-empty` | Permanently delete everything in the recycle bin |
| GET | `/api/search?path=<abs>&q=<text>` | Recursive file/folder name search under `path` (cap 500 hits) |
| POST | `/api/props` | File/folder properties `{path}` — size, times, MD5/SHA-256 |
| POST | `/api/extract-all` | Extract a whole archive into its own folder `{path}` (skip-existing, Zip-Slip safe) |
| POST | `/api/zip-create` | Pack multiple paths into one ZIP `{paths: []}` and return it as a download |
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
  "deniedPaths": [],
  "security": {
    "requireLocalOrigin": true,
    "allowedOrigins": [],
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
| `deniedPaths` | Path blacklist: listed files/folders cannot be browsed, searched, read or modified (hidden from listings too). The app's own install directory is always denied (except the app-managed recycle bin, which stays readable for its preview/download; writes still blocked) |
| `blockExecutables` | Reject executable files (default: true) |
| `security.allowedOrigins` | Webpage allowlist for API callers. The app's own UI origin (localhost/127.0.0.1 on the configured ports; defaults 20000/20443; same-origin calls with matching `Host` included) is always allowed and not subject to the list; entries are full origins (`http://192.168.1.10`: exact scheme+host+port) or bare domains (`192.168.1.10`: exact hostname, any scheme/port); exact match only - no wildcards, no subdomain inheritance (`b.a.com` needs its own entry); empty list (default) rejects every other site; `Origin: null` is always rejected |
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

- **Customizable UI** — The UI is a standalone `index.html` file. Source at `src/index.html`; deployed to `C:\html-FileNest\index.html` on install and refreshed on upgrade (the replaced file is backed up to `index.html.bak`). Edit the deployed file to personalize — restarting from the installed location never overwrites it (hot-reloads on refresh)
- **Grid / List views** — Toggle between visual grid and detailed list
- **Directory tree sidebar** — Clickable tree showing current path hierarchy and subdirectories; "All Locations" root node for quick return
- **Collapsible sidebar** — Collapse/expand the directory tree via the toolbar button or `Ctrl+B`; the choice is remembered across sessions
- **Desktop entry** — The "All Locations" page shows a "Desktop" folder entry alongside the drives, opening the user's Desktop folder like any directory
- **Filename search** — Press Enter to search file/folder names — the current folder first, then optionally continue into subfolders after a confirmation
- **Sorting** — Sort by name, size, modified time, or type (ascending/descending)
- **Grouping** — Group files by category (folders, documents, images, videos, etc.)
- **File preview** — Single-click to preview images, video, audio, PDF, text/code, Office (doc/docx/xls/xlsx/ppt/pptx), Design (PSD/PSB/AI); click archives (ZIP/JAR/TAR/GZ/RAR/7Z) to browse like a folder (navigate subdirectories, preview/download files inside); drag the left edge of the preview panel to resize
- **Adaptive layout** — Filenames wrap naturally at any zoom level; list columns scale with font size
- **Preview header** — 📂 open with system handler + ⬇ download button
- **Open with system handler** — Double-click or use the 📂 button in the preview header
- **Context menu** — Right-click for Open, Preview, Reveal, Rename, Copy to…, Move to…, Delete; archive files add "Extract here" (unpack the whole archive into its own folder); files inside an archive show Preview / Download / Open with system handler
- **Drag & drop upload** — Drag files into the window to upload
- **Keyboard shortcuts** — `Enter` (open selected), `Delete` (delete selected), `F2` (rename), `Backspace` (up), `Alt+←/→` (back/forward), `F5` (refresh), `Ctrl+B` (collapse/expand sidebar), `Ctrl+L` (path bar), `Ctrl+Shift+N` (new folder), `Ctrl+U` (upload)

## Restart / Stop the service

After changing `config.json` (ports, certificates) or replacing the program, restart the service:

- **UI**: click the ⏻ button at the right end of the toolbar and confirm
- **Command line**: `html-FileNest.exe --restart` (restart) or `html-FileNest.exe --stop` (stop)

Both work on the deployed copy: the old `--serve` process is stopped, then the service is
relaunched automatically about 2 seconds later (config re-read on startup, so port and
certificate changes take effect). No reboot or logoff needed.

## Uninstall

Run in an **elevated** command prompt:
```bash
html-FileNest.exe --uninstall
```

This removes the scheduled task, the local certificate and the UI files (`index.html`, `index.html.bak`); the exe itself and `config.json` remain in `C:\html-FileNest\` (delete the folder manually for a full cleanup).

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
- A: The request's `Origin` header is not the app's own UI origin and not registered in `security.allowedOrigins`. Register the calling site in `config.json` (full origin or bare domain - see the API Reference above); the change takes effect without restart.

## Changelog

**v2.1.0.3 (Latest)**

**New features**
- **New folder-open flow ("enter first, read once, stream in")**: opening a folder enters it immediately with a "Loading" state, reads the complete listing in one request pre-sorted by the user's chosen sort order (sorted at the source), then streams the files into the page in batches. The UI no longer uses the two-phase `fast=1` + `stat-batch` backfill — one request per folder open instead of one plus a series of backfill calls. (`GET /api/ls?fast=1` and `POST /api/stat-batch` remain available for API consumers.)
- `GET /api/ls` accepts `sort=<key>` (name / size / mtime / type) and `asc=0/1` to return the listing in the requested order — folders always first, name as tiebreaker, same semantics as the UI sort. `sort=size/mtime` implies a full stat read even with `fast=1`, so the order is always exact.

**Improvements**
- Dead network shares no longer stall the server: an unreachable UNC host is rejected instantly for 5 minutes after its first slow failure instead of repeatedly blocking the fs thread pool, and network share enumeration now has a 15-second timeout, negative caching and startup warm-up — a dead SMB mapping can no longer drag the drives page.

**Fixes**
- Background right-click now always shows the app's New / Upload / Refresh menu. Previously it fell back to the browser's context menu when the click landed on virtual-scroll spacers, group headers, empty states or grid gaps (i.e. in most large folders).
- The preview pane now closes automatically and releases its Blob memory when you navigate to a different folder; refreshing the same folder keeps it open.
- `POST /api/stat-batch` is usable again: a parameter-count defect made every call fail, and an abort-detection defect silently returned empty results on live requests.
- The desktop entry no longer disappears from the first screen after background loading (stale view cache).

**v2.1.0.2**

**New features**
- **Two-phase directory loading ("open first, stream details")**: `GET /api/ls?fast=1` returns the listing immediately with names/types only (zero per-file `stat` calls — types come from the directory read itself), and the new `POST /api/stat-batch` endpoint fills in size/modified time in batches of 1000 while the folder is already open and usable. Large folders that used to block for seconds now open in a fraction of a second; size/modified columns show "…" until backfilled, and the list re-sorts once when backfill completes. Plain `/api/ls` (no `fast`) behaves exactly as before, including the response fields.
- `POST /api/stat-batch`: read-only batched `stat` lookup (max 2000 paths per request) with the same path guards as `/api/ls`.

**Improvements**
- **Rapid navigation never stacks operations**: starting a new navigation immediately aborts every in-flight listing/backfill/search request and discards late responses from the previous one, and the server stops processing a request the moment the client disconnects (no orphaned stat loops starving the next request). Repeatedly entering and leaving large folders stays smooth instead of getting progressively laggier.
- **Virtual scrolling for large lists**: lists over 400 rows render only the visible window (~60 DOM rows regardless of folder size) instead of rebuilding a 10,000-row DOM tree on every navigation, removing accumulated GC pressure in long sessions. Smaller lists render exactly as before.
- **Bounded directory cache**: visited directories are served instantly from cache for 90 seconds (back/forward is immediate), capped at 20 directories with automatic expiry; sort results are memoized, so re-entering a folder does not re-sort thousands of items.
- Directory listing metadata lookup is now concurrent (batches of 200) instead of one-by-one serial `stat`; listing time no longer scales with per-file latency — on a 10,000-file folder this step dropped from ~700ms to ~200ms locally, far more on slow disks/UNC shares (it was ~73% of total listing time).
- The Web UI enters the folder on the first response and backfills size/modified columns progressively without blocking interaction; huge lists render through `DocumentFragment` aggregation to avoid thousands of reflows.
- Filesystem thread pool widened to 64 (`UV_THREADPOOL_SIZE=64`) in the service process and `start.vbs`, so slow or unreachable network shares can no longer monopolize the default 4-thread pool; existing `start.vbs` files are upgraded automatically on deploy.
- Back/forward navigation is instant: visited listings are cached client-side for 90 seconds (any write operation invalidates the cache), so returning to a folder shows it immediately without re-fetching or re-backfilling.
- `/api/stat-batch` per-batch limit raised to 2000 and its real-path blacklist guard now resolves the parent directory once per folder and detects symlinks via `lstat` (same junction-safe rejection semantics as before, far fewer system calls).

**v2.1.0.1**

**New features**
- **Path blacklist (`deniedPaths`) for tamper-proofing**: files and folders on the blacklist cannot be browsed, searched, read or modified — they are hidden from directory listings and search results, and every operation (preview, download, open, save, rename, move, copy, delete, pack) rejects them with 403. The app's own install directory (config, certificates, cache) is always denied and cannot be un-denied (the only exception is the app-managed recycle bin, which stays readable for the recycle UI's preview and download — writes remain blocked); `deniedPaths` in `config.json` adds more entries and takes effect on hot reload. Junction/symlink redirects into denied trees are blocked via real-path resolution.

**Fixes**
- The "New" menu no longer misses "Text Document" and "BMP Image" on Windows 11: registrations without a creation value (`Handler`-only or empty `ShellNew` keys) now appear for text/document/image perceived types, exactly like the Windows 11 menu.
- Pseudo registrations that Windows keeps out (compressed-folder leftovers, third-party entries such as FontCreator projects) no longer appear in the "New" menu.
- Menu labels for types registered by several applications (e.g. Office and WPS) now match Windows exactly by resolving the default association per user first (UserChoice, then HKCU over HKLM).
- The "New" menu now matches Windows' own right-click "New" menu exactly, item by item. The `ShellNew` registry lookup covers every registration location Windows consults - the top-level `.ext\ShellNew` key, the nested `.ext\<ProgID>\ShellNew` key (the standard registration spot for Office/WPS documents and compressed folders, e.g. `.docx\Word.Document.12\ShellNew`), the 32-bit (WOW64) registry view used by 32-bit applications such as WPS, every loaded user's `Software\Classes` hive, and the Explorer New-menu cache as a fallback - and items Windows leaves out stay out: types without a default `ProgID`, registrations without a file-creation mechanism, and handler-based registrations without a label source (e.g. Library Folder) are filtered exactly like Explorer does.
- Menu labels now read exactly like Windows': the file type's friendly name (`FriendlyTypeName`, indirect resource strings resolved), no "(.ext)" suffix, sorted the Windows way (folder, shortcut, then by extension). Previously the ShellNew `ItemName` string was shown instead, which never matches the Explorer menu text.
- New files are now created exactly the way Windows creates them, per the type's `ShellNew` registration: template copy (`FileName`), initial content (`Data`), creation command (`Command`, run with `%1` = the new file path), or empty file (`NullFile`). Previously every new file was created empty regardless of type.

**Improvements**
- "New" menu entries carry the real Windows file type icons.
- "Shortcut" (`.lnk`) is now offered in the "New" menu and creates a real shortcut (target dialog, `WScript.Shell`); the executable/`.lnk` write block stays in force for every other path.

**v2.1.0.0**

**New features**
- **Caller webpage allowlist (`security.allowedOrigins`)**: the app's own UI origin (localhost/127.0.0.1 on the configured ports - default `http://127.0.0.1:20000`, `https://127.0.0.1:20443`; same-origin calls with matching `Host` are also always allowed) is always allowed and is not subject to the whitelist. The whitelist only applies to other sites - entries may be full origins (`http://192.168.1.10`) or bare domains (`192.168.1.10`); an empty list (the default) rejects every other site. `Origin: null` is always rejected. Requests with no `Origin` (curl, local scripts) are allowed as before.
- **One-command service restart/stop for deployed installs**: CLI commands `--restart` (stop then relaunch the service) and `--stop`, endpoint `POST /api/service-restart`, and a restart button at the bottom of the left sidebar (with icon, label, confirmation and auto-reload). Service control runs through a one-shot VBS that kills only the same-exe `--serve` instance (per-install precise), so multiple installs never affect each other.

**v2.0.1.0**

**Fixes** (security)
- Upload names `.` / `..` are now rejected — `POST /api/upload` with `name=..` wrote the file into the *parent* directory of the target folder (CWE-22), letting a crafted upload escape the selected folder.
- The extension policy now covers every write path: upload, rename, move, copy, save and new-file creation all reject blocked executable extensions, including trailing-dot/space name variants such as `evil.exe.` / `evil.exe ` (Windows silently rewrites them to `evil.exe`). Previously any of these paths could plant an executable the policy was meant to block.
- The Host/Origin security gate now also protects the legacy routes (`/ping`, `/open`, `/exists`, `/reveal`, `/open-folder`, `/cache/clear`) and `/api/*`: a forged `Host: evil.example` or a cross-site `Origin` receives 403 instead of a served response with a reflected `Access-Control-Allow-Origin`. Requests carrying neither header (curl, local scripts) are unaffected.

Security fixes contributed by [@sucloudflare](https://github.com/sucloudflare) via [#1](https://github.com/FatFatYoung/html-FileNest/pull/1) — thanks!

**v2.0.0.0**

**New features**
- Collapsible sidebar — collapse/expand the directory tree from a toolbar button or `Ctrl+B`; the preference persists across sessions.
- "Desktop" entry alongside the drives on the "All Locations" page — opens the user's Desktop folder like any directory. The path is resolved by the new `/api/desktop` endpoint (redirected Desktops via the User Shell Folders registry key; configured `roots` restrictions respected).
- "New file" alongside "New folder" — toolbar button and blank-area context menu create an empty file in the current folder via the new `POST /api/mkfile` (never overwrites an existing name).
- Network shares in "All locations" — mapped network drives, remembered UNC shares and active SMB connections (remote shares on the user's machine) are listed alongside local drives with a 🌐 marker (via `net` in `GET /api/drives`, refreshed every 60s). A single product-level rule applies to every entry on every machine: any UNC path of any depth (`\\host\share\...`, from any source) normalizes to the host root `\\server`, deduplicated case-insensitively, so the list sits at the same level as `C:\`/`D:\` and browsing one entry shows all of that host's shares instead of a single visited subdirectory; mapped drive letters keep the letter as the entry.
- "My Favorites" — a favorites module below the drives on the All-Locations page plus a folder node in the sidebar; right-click any file/folder and pick "Add to Favorites" / "Remove from Favorites" (toggle), or use the ☆ button in the preview pane header. Persisted to `favorites.json` (new `GET/POST /api/favorites`, `POST /api/unfavorite`).
- Unified "New" menu — "New folder" and "New file" merge into one hover flyout menu (expands to the right on mouse-over, auto-closes when the pointer leaves) listing folder plus every file type registered in the Windows "New" context menu (live from registry `ShellNew` via `GET /api/newfile-types`); the create dialog pre-fills the matching extension and auto-appends it if removed.
- Multi-select with batch operations — `Ctrl+Click` toggles items, `Shift+Click` selects a range, `Ctrl+A` selects all, and dragging on empty space draws a rubber-band box that selects everything it touches (`Esc` or a plain click on empty space clears). With several items selected the context menu switches to batch actions: pack-and-download as one ZIP (`POST /api/zip-create`), batch rename (pattern with `{name}`/`{n}`, extension preserved, iterates `POST /api/rename`), batch copy/move into a target folder and batch delete.
- Recycle bin — deleting moves items into an app-level recycle bin instead of destroying them. A "Recycle Bin" node sits in the sidebar and beside the drives on the All-Locations page; deleted items keep their original path and deletion time and can be restored (name collisions get a " (restored)" suffix), permanently deleted one by one, or purged all at once. New `GET /api/recycle` and `POST /api/recycle-restore` / `/api/recycle-purge` / `/api/recycle-empty`; `POST /api/rm` now recycles by default and only truly deletes with `purge: true`.
- "Extract here" for archives — a banner button at the top of the archive browse view and the file context menu unpack the whole zip/jar/tar/gz/tgz/rar/7z/bz2/xz into its own folder in one step (`POST /api/extract-all`); existing files are skipped, never overwritten; entry paths are validated against Zip-Slip. zip/tar/gz are handled in-process, rar/7z/bz2/xz via 7-Zip.
- Tree search — pressing Enter in the search box matches file and folder names in the current folder first, then asks "Current folder searched. Continue into subfolders?"; confirming continues with the recursive search across the whole tree (`GET /api/search`, capped at 500 hits shown as relative paths with a truncated notice). Typing alone never triggers a search — only Enter does. Double-clicking a hit opens it; clearing the box returns to the folder.
- Properties dialog — right-click any item for name, type, location, size (contained-item count for folders), created/modified times and MD5/SHA-256 for files (`POST /api/props`).
- `Enter` opens the selected item (double-click equivalent) and `Delete` deletes it via the confirm dialog.
- The recycle bin gains batch restore-to-original-location and batch permanent-delete (sidebar and multi-select context menus).
- Language switcher — a dropdown to the right of the toolbar upload buttons, persisted in `localStorage.fn_lang`; every label, button, menu, dialog, toast, group title, tooltip and the window title switches language live.

**Improvements**
- Friendly delete confirmation — deletion asks via the in-app dialog (multiline message, Esc / click-outside cancels) instead of the browser's native confirm/alert; success shows a status toast.
- Delete button in the preview header (next to download) deletes the file currently being previewed through the same confirmation flow.
- Recycle-bin items can be previewed like normal files (single-click or context menu); the preview header's delete button permanently deletes recycled items and the favorite button is hidden for them.
- Context menu gains "Copy filename" (full name with extension, to the clipboard); file groups in the list expand/collapse when their header is clicked (▾/▸); the preview pane's open/delete buttons now use the same icons as the context menu.
- Sorting defaults to modification time descending and the chosen sort is remembered across sessions; grouping keeps folders/images/archives as fixed leading groups and breaks the remaining items out by extension (extension-less files in their own group); entry cards no longer take part in sorting/grouping. Entry-card context menus slim down to open / reveal / properties.

**Fixes**
- Shortcuts no longer fire while typing in inputs / the preview text editor, or while a dialog is open.
- UI file lifecycle. Installing or upgrading now always refreshes `index.html` to the new version's UI — previously a leftover `index.html` from an old install was never overwritten, so UI updates never reached the deployment. The replaced file is backed up to `index.html.bak` first, and launching from the installed location still leaves the UI untouched so user customizations survive restarts. `--uninstall` now also removes `index.html` and `index.html.bak` (previously they were left behind and poisoned the next install).
- "Extract here" in the file context menu now actually appears for archives — the extension test compared the dot-less value returned by the normalizer against a dotted list and never matched, so the menu item never showed for any file.
- Whole-archive extraction no longer rejects every entry with "unsafe entry path" when the archive sits at a drive root — the containment prefix gained a doubled separator at a root, so every entry failed the prefix check even with a clean name. The Zip-Slip guard is also hardened: leading separators are stripped from entry names, and `..` segments or drive-qualified names are rejected outright.

**v1.4.2.3**
- Security: archive extraction hardening (three fixes). (1) `extract7zEntry` (7z/RAR extraction) now passes `-spd -y --`: wildcard characters in entry names are matched literally and cannot redirect extraction, and entry names starting with `-` cannot be parsed as 7-Zip options. (2) TAR extraction now validates `dataOffset + size` against the buffer before slicing — a TAR header declaring an out-of-range size used to be silently truncated by `Buffer.slice`, producing a corrupted download with no error; it is now rejected as `corrupted archive: entry size out of range`. (3) Single-file GZ extraction now verifies the requested entry name against the virtual entry name — previously the `entryPath` argument was ignored and the entire decompressed content was returned for any request.
- Robustness: review pass over the remaining audit items. Archive entries: filename decoding now resolves very short names (≤5 bytes) by comparing UTF-8 vs GB18030 plausibility (fixes GBK-encoded names that happen to also be valid UTF-8 byte sequences); TAR parsing stops at the target entry instead of scanning the whole archive; `archive-extract` reuses the full MIME table (`.webm`/`.mkv`/`.flac`/Office types inside archives no longer fall back to `application/octet-stream`); entry listing sorts with a fixed `zh` locale; all 7z invocations pass `-y`; a missing 7z.exe now raises a clear error instead of a silent spawn failure.
- Robustness: temporary-file lifecycle. Upload leftovers (`*.upload.<pid>.<ts>`) left behind by crashes/power loss are now tracked via a directory journal and swept after 24h; extracted "open with system" copies under `%TEMP%\html-FileNest-open` are deleted shortly after use and swept as a fallback; in-flight upload temp files are cleaned up on process exit.
- Robustness: `/reveal` degrades to opening the containing folder when the path contains a comma (`explorer /select,` misparses it); `/api/write` accepts bodies up to 50MB (was 10MB) so large text files can be saved; config `${ENV_VAR}` expansion is restricted to a safe variable whitelist and roots must be absolute paths — drive-letter or UNC network paths (relative paths are rejected).
- UI: preview internals deduplicated — archive-in preview and normal-file preview now share a single renderer (adds one preview type in one place). Fixed a Blob memory leak: Object URLs created for PDF/AI previews are now tracked and revoked on preview switch/close (previously they accumulated for the lifetime of the page). Removed a dead unreachable preview branch.

**v1.4.2.2**
- Security: fixed HTML injection in Office document preview (P1 #3). Preview HTML (mammoth output for DOCX) was assigned to `innerHTML` unsanitized — a crafted document could execute script in the app origin. All preview HTML now passes DOMPurify 3.4.17 (vendored pinned copy); the file-tree error message (which can embed hostile filenames) is sanitized as well.
- License compliance: full third-party license texts (Bun, JSZip, mammoth, word-extractor, SheetJS, DOMPurify, 7-Zip notice) are now embedded inside the executable. Run `html-FileNest.exe --licenses` to display them at any time.
- Docs: corrected the SheetJS license note — v0.20.3 (SheetJS CDN) is Apache-2.0, same license as before; the previous "newer versions use a different license" statement was outdated.

**v1.4.2.1**
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
| [DOMPurify](https://github.com/cure53/DOMPurify) | HTML sanitization for document preview (XSS/HTML injection defense) | Apache-2.0 (elected from MPL-2.0 OR Apache-2.0) |
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
