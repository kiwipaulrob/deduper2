# Installation & Setup Guide

## Prerequisites

Before running DeDuper, ensure you have:

- **Windows Operating System** (XP SP3 or later, tested on Windows 7+)
- **iTunes** installed and functioning
- **Administrator Access** (may be required for some file operations)
- **VBScript Support** (built into Windows by default)

## Download & Installation

### Option 1: Clone from GitHub

```bash
git clone https://github.com/yourusername/deduper2.git
cd deduper2
```

### Option 2: Direct Download

Download `DeDuper.vbs` and save it to a convenient location on your computer.

## Running the Script

### Method 1: Double-click (Recommended)

1. Navigate to the script location
2. Right-click `DeDuper.vbs`
3. Select "Open with" → "Windows Cscript"
   - If this option doesn't appear, select "Choose another app" → "More apps" → browse to:
   - `C:\Windows\System32\cscript.exe`

### Method 2: Command Line

1. Open Command Prompt (cmd.exe) or PowerShell
2. Navigate to the script directory
3. Run one of these commands:

```batch
cscript DeDuper.vbs
```

Or:

```batch
wscript DeDuper.vbs
```

(Note: `cscript` shows progress bar; `wscript` uses graphical dialogs)

### Method 3: Create a Shortcut

For easier access, create a Windows shortcut:

1. Right-click on the `DeDuper.vbs` file
2. Select "Create shortcut"
3. Right-click the new shortcut → "Properties"
4. Set the target to:
   ```
   C:\Windows\System32\cscript.exe "C:\full\path\to\DeDuper.vbs"
   ```
5. Optionally set the "Run" option to "Minimized"
6. Click "OK" and save

## Initial Setup

### First Run Checklist

- [ ] iTunes is running and not actively playing
- [ ] You have a backup of your iTunes library
- [ ] You have identified which tracks to process (entire library or specific playlist)
- [ ] You understand the duplicate-handling options you want to use

### Recommended Settings for First Run

For safety, we recommend these settings when first using DeDuper:

```vbscript
Uncheck=True            ' Marks duplicates as unchecked for review
Testing=False           ' Actually removes duplicates
Intro=True              ' Show all introduction dialogs
Outro=True              ' Show summary report
```

## Configuration

Edit `DeDuper.vbs` with a text editor (Notepad, Notepad++, VSCode, etc.) to adjust settings:

1. Open the file with your preferred text editor
2. Locate the "Initialise variables" section (around line 170)
3. Modify the settings as needed
4. Save the file
5. Run the script with your new settings

### Common Configuration Scenarios

#### Scenario 1: Preview Mode (Safest for First Run)

```vbscript
Testing=True            ' No actual deletion
Uncheck=True            ' Mark potential duplicates as unchecked
Intro=True              ' Show introduction
Outro=True              ' Show summary
```

#### Scenario 2: Remove Logical Duplicates Only

```vbscript
CheckTypes=True         ' Select types manually
SameSongs=False         ' Don't look for same songs
KeepLarge=True
Uncheck=True
```

#### Scenario 3: Archive Instead of Delete

```vbscript
Archive=True            ' Copies to Archive folder
Uncheck=True            ' Review before finalizing
Testing=False
```

#### Scenario 4: Prefer Specific Format

```vbscript
PrefKind=True
PrefKinds=".mp3.m4a.flac"   ' MP3 > M4A > FLAC
Uncheck=True
```

## Troubleshooting

### "Permission Denied" Error

**Solution:** Run the script as Administrator
1. Right-click Command Prompt or PowerShell
2. Select "Run as administrator"
3. Navigate to the script location
4. Run: `cscript DeDuper.vbs`

### "iTunes Not Found" Error

**Solution:** Ensure iTunes is properly installed and running
1. Open iTunes manually first
2. Let it fully load
3. Then run DeDuper

### Script Takes Too Long

**Solution:** Disable progress bar during scan phase
```vbscript
ProgScan=False          ' Disable progress bar while scanning
```

### Tracks Not Being Removed

**Possible causes:**
- Testing mode is enabled (`Testing=True`)
- Tracks are in the cloud/streaming service
- Selected tracks aren't actually duplicates

**Debugging:**
```vbscript
Testing=True            ' Enable test mode first
Uncheck=True            ' See what WOULD be removed
Tracing=True            ' Show detailed messages
```

### Script Crashes or Hangs

If the script crashes with a specific error:

1. Note the error message
2. Check the [GitHub Issues](https://github.com/yourusername/deduper2/issues)
3. Create a new issue with:
   - Windows version
   - iTunes version
   - Error message/behavior
   - Number of tracks in library

## Backup Your Library

Before running DeDuper on your full library:

1. **Back up iTunes library:**
   - Close iTunes
   - Copy entire iTunes folder (usually in `%USERPROFILE%\Music\iTunes`)
   - Paste to external drive or cloud storage

2. **Export playlists:**
   - In iTunes: File → Library → Export As Playlist
   - Save to safe location

3. **Test on a copy first:**
   - Use backup to test DeDuper on a subset of tracks

## After Running

### Review the Report

After completion, examine the summary report carefully:
- Number of duplicates found and removed
- Statistics on play count merging
- Any tracks that couldn't be removed
- Total execution time

### Verify Results

1. Open iTunes and check affected playlists
2. Confirm tracks still appear in expected locations
3. Verify play statistics were merged correctly
4. Check that no important tracks were accidentally removed

### If Problems Occur

1. Restore from backup
2. Try again with different settings
3. Run in test mode (`Testing=True`) first
4. Check that `Uncheck=True` to review before removing

## Getting Help

- **Original Script:** [samsoft.org.uk](http://samsoft.org.uk/iTunes/scripts.asp)
- **Bug Reports:** Create an issue on GitHub
- **Feature Requests:** Discuss on GitHub Discussions

---

**Happy deduping!** Always backup first. 🎵
