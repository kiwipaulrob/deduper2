# Development Progress

## Overview

**DeDuper** is a Windows VBScript utility that detects and removes duplicate tracks from iTunes libraries. It connects to iTunes via COM automation, scans the user's library through six sequential passes, and supports multiple strategies for selecting which duplicate to keep.

### Core Facts

| Attribute         | Detail                                      |
|-------------------|---------------------------------------------|
| **Language**      | VBScript (`.vbs`), ~2000 lines              |
| **Current Version** | 1.0.5.1 (May 2026)                        |
| **Author**        | Steve MacGuire                              |
| **License**       | GNU General Public License v3.0             |
| **iTunes Integration** | COM automation (`iTunes.Application`) |
| **Target OS**     | Windows 7+ (iTunes required)                |

---

## Architecture

### Six Scan Passes

The script processes the user's iTunes selection through six sequential scan passes, each targeting a specific type of duplicate:

| Pass # | Constant            | What It Detects                                        |
|--------|---------------------|--------------------------------------------------------|
| 0      | `PASS_PLAYLIST`     | Identical tracks within the same playlist (same `PersistentID`) |
| 1      | `PASS_LOGICAL`      | Multiple iTunes entries pointing to the same physical file path |
| 2      | `PASS_PHYSICAL`     | Different files with identical metadata *and* file size |
| 3      | `PASS_ALTERNATE`    | Different files with matching metadata but different file sizes |
| 4      | `PASS_MISSING`      | Metadata matches but the physical file no longer exists |
| 5      | `PASS_SAME_SONG`    | Same artist + song name across different albums        |

### Six Duplicate Types

The six passes map directly to six `DUPE_*` constants:

| Constant              | Description                                    |
|-----------------------|------------------------------------------------|
| `DUPE_LOGICAL`        | Same file path appears more than once          |
| `DUPE_PHYSICAL`       | Same metadata + size, different physical files |
| `DUPE_ALTERNATE`      | Same metadata, different file size             |
| `DUPE_MISSING`        | Same metadata but file is missing              |
| `DUPE_SAME_SONG`      | Same artist + name, different album            |
| `DUPE_PSEUDO_PHYS`    | Physical dupe detected via alternate path      |

### Deletion Strategies

| Strategy                    | Description                                               |
|-----------------------------|-----------------------------------------------------------|
| Keep Largest File           | Retain the highest-quality version                        |
| Keep Smallest File          | Minimize storage usage                                    |
| Keep by Age                 | Oldest/Newest date added, or Oldest/Newest date modified  |
| Keep by File Type           | Prefer certain formats (e.g., `.mp3` over `.m4a`)         |
| Keep by Rating              | Prioritize higher-rated tracks                            |
| Keep Checked Track          | Manual preference when one duplicate is checked           |

### Data Preservation

- **Playlist membership** — tracks are re-added to all playlists after replacement
- **Play statistics** — play counts, skip counts, and dates are merged
- **Disc number handling** — disc 0 treated as disc 1 (configurable)
- **Link swapping** — alternate dupes are swapped to preserve "date added"

### File Handling

- **Recycle Bin** — default; files sent to Windows Recycle Bin
- **Direct deletion** — optional bypass of Recycle Bin
- **Archive** — files moved to `<Media Folder>\Archive` instead of deleted
- **Empty folder cleanup** — folders emptied by deletion are automatically removed
- **Network paths** — supports UNC paths and network-mounted media

---

## Current Status (v1.0.5.1)

### Completed in v1.0.5.1

- ✅ **Named constants** — `DUPE_*` (6 types), `PASS_*` (6 scan passes), `KIND_*` (file/stream) replace all magic numbers in scan logic
- ✅ **AppConfig class** — scaffolding class encapsulates all user-configurable settings (`KeepLarge`, `PrefAge`, `PrefKind`, `UseTrash`, `Testing`, `Disc0as1`, etc.) in one place
- ✅ **Module-level RegExp engine** — single `RE` object initialized at startup; eliminates COM object allocation per `Despace()` call
- ✅ **GetMediaPath fix** — removed duplicate `P=Left(L,Instr(L,A)-2)` assignment that could raise an unhandled runtime error outside the error guard
- ✅ **Recycle error scope fix** — message-building and `MsgBox` moved outside `On Error Resume Next` block via flag variable
- ✅ **Duplicate `Debug=True` removal** — cleaned up copy-paste leftover
- ✅ **COM caching** — `T.Kind` read once into `tKind` before `With T` block; reduced from 3 COM boundary crossings to 1 per track per pass

### Known Issues

- ⚠️ **AppConfig class** — scaffold only; global variables remain authoritative. Full migration pending.
- ⚠️ **Error handling** — some `On Error Resume Next` blocks remain broader than ideal, particularly in file-deletion paths
- ⚠️ **Performance** — linear O(n) per pass against the full track selection. Large libraries (>50K tracks) show noticeable scanning time
- ⚠️ **No test suite** — all testing is manual on Windows with a live iTunes installation
- ⚠️ **Windows-only** — no cross-platform support (bound to iTunes COM and Windows Script Host)

### Test Coverage

Currently **no automated tests exist.** Testing relies on:
- Manual cscript execution
- Test Mode (`Testing=True`) to preview deletions
- User-submitted bug reports from Apple Support Communities threads

---

## Future Ideas

### Short-term (v1.0.6.x)

| Idea | Description | Priority |
|------|-------------|----------|
| **Complete AppConfig migration** | Move all config reads from globals to `AppConfig` instance; deprecate global config variables | High |
| **Narrow error guards** | Audit all `On Error Resume Next` blocks; push string formatting and dialogs outside guards | High |
| **Progress bar granularity** | Show per-pass progress (currently shows overall only) | Medium |
| **Same-song skip/playlist merge** | Improve playlist handling for `PASS_SAME_SONG` — preserve statistics from both original tracks | Medium |
| **Batch uncheck refinement** | When `Uncheck=True`, group unchecked duplicates more intuitively in iTunes UI | Low |

### Medium-term (v1.1.0)

| Idea | Description | Priority |
|------|-------------|----------|
| **Unit test harness** | Write a VBScript test runner (or move core logic to a testable stub) | High |
| **CI pipeline** | GitHub Actions workflow to lint VBScript (syntax check via cscript on a Windows runner) | Medium |
| **Dry-run report export** | Export Test Mode results to CSV/HTML instead of just MsgBox | Medium |
| **iTunes library XML export** | Option to process an iTunes Library.xml export file instead of live COM — enables headless/review use | Medium |
| **Config file support** | Load settings from an `.ini` or JSON file instead of editing the script | Medium |
| **Multi-language UI** | Extend existing Dutch support framework to additional languages | Low |

### Long-term (v2.0+)

| Idea | Description | Priority |
|------|-------------|----------|
| **MusicBee / MediaMonkey support** | Abstract the media library interface to support non-iTunes players | High |
| **Port to PowerShell** | Rewrite as a PowerShell module for modern Windows scripting, better error handling, and native unit testing (Pester) | Medium |
| **Cross-platform port** | Port to Python for macOS/Linux support (Music.app, Rhythmbox, etc.) | Low |
| **CLI mode** | Headless operation with command-line arguments for batch/automation use | Low |

---

## Changelog Highlights

For the full version history, see [CHANGELOG.md](./CHANGELOG.md).

| Version | Date | Key Changes |
|---------|------|-------------|
| 1.0.5.1 | 2026-05-22 | Named constants, AppConfig class, RegExp caching, bug fixes |
| 1.0.4.10 | 2025-05-16 | Added error handling for edge cases |
| 1.0.4.3 | 2024-09-15 | Same-song duplicate detection |
| 1.0.4.1 | 2024-08-10 | Keep by date added/modified |
| 1.0.3.2 | 2024-03-10 | Archive to folder option |
| 1.0.1.9 | 2023-02-10 | File type preference |
| 1.0.0.1 | 2022-01-10 | Initial release |

---

*Last updated: June 2026*
