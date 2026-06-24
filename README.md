# SharePoint Bulk Downloader

A browser extension for **Chrome** and **Firefox** that downloads hundreds of files from a SharePoint folder automatically — one by one, directly to a local folder, with no "Save As" dialog per file.

Built to solve the problem of downloading large SharePoint folders (96 GB+, 600+ files) where downloading everything at once causes crashes, but clicking each file manually would take hours.

---

## Features

- **No "Save As" dialogs** — Uses the File System Access API: you pick a destination folder once, files stream directly to disk.
- **Sequential downloads** — One file at a time, preventing the browser crash that bulk downloads cause.
- **Live statistics** — Current speed, ETA, elapsed time, bytes downloaded, average speed, file count.
- **Per-file progress** — Byte-level progress bar for each file while it downloads.
- **Completed files list** — Full log of every downloaded file with status (✓ ok / ✗ failed / – skipped) and timestamp.
- **Activity log** — Color-coded, timestamped log of every event. Exportable as `.txt`.
- **Retry failed files** — Re-queue only failed files without re-downloading successful ones.
- **Skip existing** — Resume interrupted sessions without re-downloading files already on disk.
- **Recursive subfolders** — Optionally mirrors the full folder hierarchy.
- **Shared folders** — Works with SharePoint folders shared with you by others.

---

## Requirements

| Browser | Minimum version |
|---------|----------------|
| Chrome / Edge | 86+ |
| Firefox | 111+ |

> The File System Access API (`showDirectoryPicker`) is required. It is available in the versions listed above.

---

## Installation

### Chrome / Edge

1. Download or clone this repository.
2. Generate the icons (see [Icons](#icons) below).
3. Open `chrome://extensions`.
4. Enable **Developer mode** (top right toggle).
5. Click **Load unpacked** and select the `sharepoint-downloader` folder.

### Firefox

1. Open `about:debugging#/runtime/this-firefox`.
2. Click **Load Temporary Add-on**.
3. Select the `manifest.json` file inside the `sharepoint-downloader` folder.

> **Note:** Temporary add-ons in Firefox are removed on browser restart. For a persistent install, the extension must be signed via [AMO](https://addons.mozilla.org/).

---

## Icons

The repository does not include binary PNG icons. To generate them:

1. Open `generate_icons.html` in any browser (double-click the file).
2. Click **Gerar e Baixar Ícones** — three PNG files download automatically.
3. Move `icon16.png`, `icon48.png`, and `icon128.png` into the `icons/` folder.

---

## Usage

1. Navigate to the SharePoint folder you want to download.
2. Click the extension icon in the browser toolbar — a new tab opens.
3. The extension auto-detects the SharePoint tab and folder path.
4. Click **Escolher pasta...** to select your local destination folder (done once per session).
5. Optionally set a subfolder name and configure options.
6. Click **Listar Arquivos** — the extension calls the SharePoint REST API to enumerate all files.
7. Review the file list, then click **Iniciar Downloads**.

Files stream directly to your chosen folder. Keep the tab open while downloads run.

### Tabs

| Tab | Description |
|-----|-------------|
| **Fila** | Files scanned and waiting to download |
| **Concluídos** | Files already downloaded, with status and timestamp |
| **Log** | Full activity log, exportable as `.txt` |

### Controls

| Button | Action |
|--------|--------|
| **Pausar** | Finishes the current file, then pauses |
| **Continuar** | Resumes from where it paused |
| **Cancelar** | Stops immediately (current file may be incomplete) |
| **Tentar novamente** | Re-queues only failed files |
| **Exportar .txt** | Saves the full session log to a text file |

---

## How it works

### File listing

The extension uses the SharePoint REST API with your existing browser session (no separate authentication needed):

```
GET /_api/web/GetFolderByServerRelativeUrl('{path}')/Files
GET /_api/web/GetFolderByServerRelativeUrl('{path}')/Folders
```

Results are paginated automatically for large folders.

### File download

Each file is fetched via `fetch()` with `credentials: 'include'` (uses your browser session cookies), then written to disk using the [File System Access API](https://developer.mozilla.org/en-US/docs/Web/API/File_System_API):

```
response.body  →  TransformStream (progress tracking)  →  FileSystemWritableFileStream (disk)
```

This pipeline streams data directly without buffering the entire file in memory — safe even for very large files.

### Speed calculation

Speed is computed as a 5-second rolling average over the bytes received from the network stream, updated every 500 ms.

---

## Troubleshooting

### "Pasta não detectada"

The extension could not read the current folder path from the URL. Use the **Caminho manual** field: paste the SharePoint folder URL or the server-relative path (`/sites/SiteName/Library/Folder`).

### Files not listing (403 / empty result)

- Make sure you are authenticated to SharePoint in the same browser session.
- For shared folders from another tenant, navigate directly to the shared link first, then open the extension.
- Check the **Log** tab for the exact API error.

### Downloads stop after browser sleeps

Chrome and Firefox may suspend tabs when the system sleeps. Keep the machine awake during large downloads, or use your OS's power settings to prevent sleep.

---

## Project structure

```
sharepoint-downloader/
├── manifest.json          # Extension manifest (Manifest V3)
├── background.js          # Service worker — opens the app tab on icon click
├── content.js             # Content script — SharePoint REST API calls
├── app.html               # Main UI (full-page tab)
├── app.js                 # UI logic, download engine, speed tracking, log
├── app.css                # Styles
├── generate_icons.html    # Utility to generate icon PNG files
├── icons/                 # Extension icons (user-generated, not tracked)
├── CHANGELOG.md
├── LICENSE
└── README.md
```

---

## Contributing

Pull requests are welcome. For significant changes, please open an issue first to discuss what you'd like to change.

---

## License

[MIT](LICENSE)
