# Changelog

All notable changes to DeDuper are documented in this file.

## [1.0.5.1] - 2026-05-22

### Added
- **Named constants** — `DUPE_*` (6 types), `PASS_*` (6 scan passes), and `KIND_*` (file/stream) replace
  all magic numbers throughout the scan logic, making intent self-documenting
- **AppConfig class** — scaffolding class encapsulates all user-configurable settings
  (`KeepLarge`, `PrefAge`, `PrefKind`, `UseTrash`, `Testing`, `Disc0as1`, etc.) in one place;
  global variables remain authoritative for now pending full migration
- **Module-level RegExp engine** — single `RE` object initialised at startup; eliminates a new
  `RegExp` COM object being allocated on every `Despace()` call (called once per track, per pass)

### Fixed
- **GetMediaPath** — removed duplicate `P=Left(L,Instr(L,A)-2)` assignment that executed
  *outside* the `On Error Resume Next` guard; on a malformed path this would raise an
  unhandled runtime error after the guarded attempt had already captured the failure
- **Recycle (error scope)** — message-building and `MsgBox` call moved outside the
  `On Error Resume Next` block using a flag variable; previously, string-formatting operations
  inside the error scope could silently swallow their own errors
- **Duplicate `Debug=True` assignment** — removed a copy-paste leftover that redundantly
  re-assigned `Debug=True` six lines after its first assignment

### Changed
- **DedupeTracks scan loop** — `T.Kind` is now read once into `tKind` before the inner `With T`
  block, reducing the loop from three COM boundary crossings per track (for `Kind`) to one;
  magic literals `1` and `3` replaced with `KIND_FILE` and `KIND_STREAM` constants

## [1.0.4.10] - 2025-05-16

### Fixed
- Error handling for additional edge cases in file path handling
- Improved robustness when detecting mixed path formats

## [1.0.4.9] - 2025-01-20

### Fixed
- Additional error traps for reported issues in duplicate detection
- Improved handling of certain edge cases in file system operations

## [1.0.4.8] - 2024-11-15

### Fixed
- Error handling improvements for various edge cases
- Better detection of pseudo-physical duplicates

## [1.0.4.7] - 2024-10-10

### Fixed
- Remove repeat spaces from track metadata before comparing
- Trim whitespace from values before comparison

## [1.0.4.6] - 2024-09-25

### Fixed
- Edge case fix in GetMediaPath function
- Improved path resolution for network locations

## [1.0.4.5] - 2024-09-23

### Fixed
- Modified Recycle function for better compatibility with network paths
- Improved network file deletion handling

## [1.0.4.4] - 2024-09-20

### Optimized
- Optimized same song duplicate playlist removal code
- Improved performance when removing duplicates from playlists

## [1.0.4.3] - 2024-09-15

### Added
- **Same Song Duplicate Detection:** New ability to detect and remove tracks with identical song name/artist but different metadata
- Automatic unchecking of duplicate songs from regular playlists
- New configuration options for same-song deduplication

### Improved
- Enhanced alternate version detection rules
- Fixed selection reset issue that occurred after deleting from playlist

## [1.0.4.1] - 2024-08-10

### Added
- **Keep by Age:** New options to preserve newest or oldest file by date added or date modified
- **Enhanced Start Screen:** Improved introduction dialog showing all selected options
- **Detailed Report:** Updated report display to show active options

## [1.0.3.6] - 2024-07-05

### Added
- Stream matching capability
- Ability to dedupe stream content (may have limitations with smallest file preference)

## [1.0.3.5] - 2024-06-20

### Added
- Swap links between alternate dupes to preserve "date added" property
- Uncheck duplicates that would be removed for manual review before deletion

### Improved
- Better preservation of iTunes metadata during duplicate removal

## [1.0.3.4] - 2024-05-15

### Improved
- Enhanced dialog boxes when archiving files
- Added troubleshooting code (can be enabled with simple edit)

## [1.0.3.3] - 2024-04-20

### Fixed
- Attempt to fix sporadic "missing object" error

## [1.0.3.2] - 2024-03-10

### Added
- **Archive Option:** New ability to archive deleted files to `<Media Folder>\Archive` instead of permanently deleting them
- Alternative to Recycle Bin deletion

## [1.0.3.1] - 2024-02-15

### Fixed
- Missing Ext function issue

## [1.0.2.9] - 2023-11-20

### Improved
- Additional error handling throughout the script
- Better exception management for file operations

## [1.0.2.8] - 2023-10-25

### Fixed
- Fixed error in track confirmation code
- Improved user interface responsiveness

## [1.0.2.7] - 2023-09-30

### Added
- Error handler for file deletion routine
- More robust file removal process

## [1.0.2.6] - 2023-08-15

### Improved
- Minor update to common code base
- Added trap for possible error in Merge routine

## [1.0.2.5] - 2023-07-20

### Added
- **Pseudo-Physical Duplicate Detection:** Detect pseudo-physical duplicates
- **Media Folder Preference:** Prefer files inside media folder when duplicates found

## [1.0.2.4] - 2023-06-10

### Added
- **Disc 0/1 Equivalence:** Option to treat disc 0 as equivalent to disc 1 for matching purposes
- Improved Recycle function for non-English Windows installations

## [1.0.2.3] - 2023-05-05

### Added
- **Deletion Mode Selection:** Option to attempt trash/Recycle Bin deletion or delete directly

## [1.0.2.2] - 2023-04-15

### Improved
- Enhanced UAC/LUA detection for scanning phase
- Better handling of elevated permissions

## [1.0.2.1] - 2023-03-20

### Added
- **Automatic Missing Duplicate Removal:** Automatic deletion of tracks marked as "missing" that are also duplicates

## [1.0.1.9] - 2023-02-10

### Added
- **File Type Preference:** Option to prefer certain file types over others when handling duplicates with different sizes
- Extensible format preference system

## [1.0.1.8] - 2023-01-15

### Added
- Dutch language support
- Framework for easy extension to other languages

## [1.0.1.7] - 2022-12-20

### Improved
- Minor update to file delete code
- Enhanced bug hunting and error detection

## [1.0.1.6] - 2022-11-10

### Added
- **Keep Smallest File:** Option to keep smallest file instead of largest
- **Repeat Prevention:** Prevent processing the same file twice in a playlist
- **Repeat Cleaning:** Clean repeats functionality

## [1.0.1.5] - 2022-10-15

### Fixed
- Potential bug when running script under UAC (User Account Control)
- Improved permission handling

## [1.0.1.4] - 2022-09-20

### Added
- **Extended Deletion Strategy:** Optional removal of alternate dupes with choice to keep largest file

## [1.0.1.3] - 2022-08-15

### Added
- **Playlist Preservation:** Feature to ensure playlist membership is preserved during duplicate removal
- Tracks added to any playlist containing removed duplicates

## [1.0.1.2] - 2022-07-10

### Improved
- Updated recycle routine for Vista and Windows 7 compatibility
- Better Windows version detection

## [1.0.1.1] - 2022-06-05

### Improved
- Updated to common code base with progress bar
- Unified progress tracking across script

## [1.0.0.6] - 2022-05-15

### Improved
- Update detection routines to handle Mac/Linux style file paths imported to database
- Rewrite recycle routine for clarity
- Better path normalization

## [1.0.0.5] - 2022-04-20

### Improved
- Enhanced distinction between logical and physical duplicates
- Automatic removal of folders emptied by duplicate deletion
- Better cleanup logic

## [1.0.0.4] - 2022-03-10

### Fixed
- Fixed issue when more than two physical copies of a file existed
- Improved duplicate grouping logic

## [1.0.0.3] - 2022-02-05

### Added
- **Extended Duplicate Detection:** Now deals with files that have duplicate metadata and same file size
- **Recycle Bin Integration:** Deleted files go to Recycle Bin instead of permanent deletion

## [1.0.0.2] - 2022-01-20

### Fixed
- Fixed race condition in processing logic

## [1.0.0.1] - 2022-01-10

### Initial Release
- Initial version of DeDuper script
- Basic duplicate detection and removal
- Support for both logical and metadata duplicates
- iTunes library integration

---

## Format Notes

- Dates are in YYYY-MM-DD format
- Versions follow semantic versioning: MAJOR.MINOR.PATCH.BUILD
- [Original changelog from DeDuper.vbs comments](http://samsoft.org.uk/iTunes/scripts.asp)
