# DeDuper

**Version:** 1.0.4.10 (May 16, 2025)  
**Author:** Steve MacGuire  
**License:** GNU General Public License v3.0

## Overview

**DeDuper** is a powerful Windows VBScript utility designed to identify and remove duplicate tracks from iTunes libraries. It handles both logical duplicates (multiple iTunes entries for the same physical file) and physical duplicates (different files with identical metadata).

The script is highly configurable and provides multiple options for handling different types of duplicates with user-friendly confirmation dialogs and progress tracking.

## Features

### Duplicate Detection
- **Logical Duplicates:** Removes multiple iTunes database entries pointing to the same physical audio file
- **Metadata Duplicates:** Identifies tracks with identical metadata (artist, album, title, etc.) even if they're different file sizes
- **Same Song Detection:** Optional detection of duplicate entries for the same song (can ignore brackets/parentheses)
- **Stream Matching:** Ability to match and dedupe streaming content

### Deletion Strategies
When handling duplicate files of different sizes, the script offers several options:

- **Keep Largest File:** Retains the highest quality version
- **Keep Smallest File:** Minimizes storage usage
- **Keep by Age:** 
  - Oldest/Newest date added
  - Oldest/Newest date modified
- **Keep by File Type:** Prefer certain audio formats (e.g., `.mp3` over `.m4a`)
- **Keep by Rating:** Prioritize tracks with higher ratings
- **Keep Checked Track:** Manual preference selection when only one duplicate is checked

### Data Preservation
- **Playlist Membership:** Automatically preserves track membership in all playlists
- **Play Statistics:** Merges play counts, skip counts, and most recent date statistics
- **Disc Number Handling:** Option to treat disc 0 as equivalent to disc 1

### File Handling
- **Safe Deletion:** Sends deleted files to Windows Recycle Bin (or direct deletion option)
- **Archive Option:** Archive deleted files to `<Media Folder>\Archive` instead of deleting
- **Empty Folder Cleanup:** Automatically removes empty directories created by deletion
- **Network Path Support:** Works with network-stored media

### User Interface
- **Progress Bar:** Visual feedback during scanning and processing phases
- **Track-by-Track Confirmation:** Optional confirmation dialog for each duplicate removal
- **Pre-process Report:** Shows selected options and processing details
- **Summary Report:** Detailed statistics including timing information
- **Test Mode:** Preview what would be removed without actual deletion
- **Uncheck Mode:** Marks potential duplicates as unchecked for manual review

## System Requirements

- Windows operating system
- iTunes installed
- Administrator/UAC privileges (for some operations)
- VBScript support (Windows built-in)

## Usage

### Basic Steps

1. **Run the Script:**
   ```batch
   cscript DeDuper.vbs
   ```

2. **Select Tracks:**
   - Select specific tracks in iTunes, or
   - Run on the entire library (using "Library" as source), or
   - Run on a specific named playlist

3. **Confirm Options:**
   - Review the presented options for duplicate handling
   - Choose which types of duplicates to process
   - Select your preferred strategy for keeping/removing files

4. **Review & Confirm:**
   - Examine the detailed report showing what will be removed
   - Confirm track-by-track removals (if enabled)
   - Review the summary report

### Configuration

The script includes configurable variables at the top of the file:

```vbscript
CheckTypes=True         ' Confirm types of dupes to remove
SameSongs=False         ' Include same-song duplicate detection
IgnoreBraces=False      ' Ignore () [] when comparing songs
MaxChars=0              ' Truncate long titles (0 = no limit)
KeepLarge=True          ' True = Keep largest file, False = Keep smallest
PrefKind=True           ' Prefer certain file types
PrefKinds=".mp3.m4a"    ' File extensions in preference order
PrefAge=0               ' 1=Oldest Added, 2=Oldest Modified, 3=Newest Added, 4=Newest Modified
PrefRating=False        ' Prefer highest rated tracks
UseTrash=True           ' Send to Recycle Bin (True) or delete directly (False)
Disc0as1=True           ' Treat disc 0 as disc 1
Archive=False           ' Archive instead of delete
Uncheck=True            ' Uncheck potential discards for review
Testing=False           ' Preview mode (no actual deletion)
```

## Examples

### Remove logical and metadata duplicates, keeping largest files
```vbscript
CheckTypes=True
SameSongs=False
KeepLarge=True
Uncheck=True
```

### Keep the oldest version of duplicate files
```vbscript
PrefAge=1               ' Oldest Date Added
Uncheck=True
Testing=False
```

### Archive duplicates instead of deleting
```vbscript
Archive=True
Uncheck=True            ' Review before finalizing
```

### Prefer MP3 over M4A format
```vbscript
PrefKind=True
PrefKinds=".mp3.m4a"
```

## Related Scripts

The DeDuper script family includes:
- **DeDuper** (this script) - Full-featured deduplication
- **DeDupeKeepSmall** - Variant that keeps smallest files
- **Duplicates** - Detect and list duplicates
- **DuplicateSong&Artist** - Find tracks with same song and artist
- **ExactDuplicates** - Find exact file duplicates

## Version History

Key improvements in recent versions:
- **v1.0.4.10** (May 2025): Additional error handling for edge cases
- **v1.0.4.3** - Same song duplicate detection and removal
- **v1.0.4.1** - Keep by date added/modified, improved UI
- **v1.0.3.2** - Archive to folder option instead of permanent deletion
- **v1.0.3.5** - Preserve date added via link swapping, uncheck for review
- **v1.0.1.9** - File type preference for duplicate selection
- **v1.0.1.6** - Keep smallest file, prevent double processing, clean repeats

## License

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

See [GNU General Public License v3.0](http://www.gnu.org/licenses/gpl-3.0-standalone.html) for more details.

## Original Author

**Steve MacGuire**  
Website: [samsoft.org.uk](http://samsoft.org.uk/iTunes/scripts.asp)

## Disclaimer

This program is distributed in the hope that it will be useful, but **without any warranty**; without even the implied warranty of **merchantability** or **fitness for a particular purpose**. 

⚠️ **Important:** Always test with a backup of your iTunes library. While the script attempts to safely remove duplicates, data loss is possible. Use the "Test Mode" and "Uncheck" options to preview changes before confirming deletion.

## Support & Updates

For the latest version, bug reports, and feature requests, visit:
- [Original Script Page](http://samsoft.org.uk/iTunes/scripts.asp)

## Contributing

This is a maintained archive of Steve MacGuire's DeDuper script. The original source and development history is maintained at the link above.

---

**Happy deduping!** 🎵
