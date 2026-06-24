# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] - 2025-06-24

### Changed
- **Architecture**: Replaced popup with a full-page browser tab (`app.html`), required for long-running downloads and File System Access API usage.
- **Download engine**: Replaced `chrome.downloads` API with the **File System Access API** (`showDirectoryPicker` + `FileSystemDirectoryHandle`). Files are now streamed directly to disk via `ReadableStream.pipeTo()` — eliminating all "Save As" dialogs.

### Added
- **Speed indicator**: Rolling 5-second window speed calculation updated every 500 ms.
- **Statistics grid**: Live display of current speed, ETA, elapsed time, bytes downloaded, file count, and average speed.
- **Per-file progress bar**: Byte-level progress for each individual file during download.
- **Completed files tab**: Full list of downloaded files with status icons (✓ ok / ✗ failed / – skipped), file size, and timestamp.
- **Activity log tab**: Dark-themed, color-coded log (INFO / SUCCESS / WARN / ERROR) with timestamps. Includes "Export .txt" button to save the full session log.
- **Retry failed**: Button to re-queue and retry only the files that failed, without re-downloading successful ones.
- **Skip existing files**: Option to skip files that already exist in the destination folder (useful for resuming interrupted sessions).

### Fixed
- API calls now use the correct SharePoint site collection URL (`_api/web` relative to `/sites/SiteName`) instead of the root origin, fixing empty results for shared folders.
- Improved SharePoint URL detection to handle more patterns including `_layouts/15/onedrive.aspx` and direct folder paths.

## [1.0.0] - 2025-06-24

### Added
- Initial release: browser extension for Chrome and Firefox (Manifest V3).
- SharePoint REST API integration to list folder contents recursively.
- Sequential downloads via `chrome.downloads` API.
- Pause, resume, and cancel controls.
- Persistent download queue via `chrome.storage.local` (survives service worker restarts).
- Automatic SharePoint folder path detection from URL parameters.
- Manual path override for cases where auto-detection fails.
- Diagnostic panel showing detected site URL and folder path.
