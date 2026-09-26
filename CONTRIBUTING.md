# Contributing to DeDuper

Thank you for your interest in contributing to **DeDuper**! This is a VBScript-based duplicate file detection and removal tool for iTunes libraries. All contributions — bug fixes, features, documentation, and testing — are welcome.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Environment](#development-environment)
- [Project Structure](#project-structure)
- [Coding Standards](#coding-standards)
- [Naming Conventions](#naming-conventions)
- [Constants and Magic Values](#constants-and-magic-values)
- [Error Handling](#error-handling)
- [Testing](#testing)
- [Pull Request Process](#pull-request-process)
- [Commit Conventions](#commit-conventions)

## Code of Conduct

This project is maintained in the spirit of constructive collaboration. Treat all contributors with respect. Harassment, personal attacks, and unprofessional conduct will not be tolerated.

## Getting Started

1. **Fork** the repository.
2. **Clone** your fork:
   ```bash
   git clone https://github.com/your-username/deduper2.git
   cd deduper2
   ```
3. Create a **topic branch**:
   ```bash
   git checkout -b feat/my-feature
   ```
4. Make your changes to `DeDuper.vbs` or the documentation files.
5. **Test** your changes (see [Testing](#testing) below).
6. Commit and push, then open a **Pull Request**.

## Development Environment

DeDuper is a **VBScript** (`.vbs`) script intended to run on **Windows** via `cscript.exe`:

- **OS:** Windows 7 or later (Windows 10/11 recommended)
- **Requirements:**
  - iTunes installed (for iTunes COM interaction)
  - VBScript support (built into Windows)
  - Administrator/UAC privileges (for some file operations)

### Local Testing on Windows

```batch
cscript DeDuper.vbs
```

### Code Review (any platform)

The script can be reviewed on any platform with a text editor. Use:
- **VSCode** with a VBScript extension for syntax highlighting
- **Notepad++** with VBScript language support
- Manual `Option Explicit` checking

## Project Structure

```
deduper2/
├── DeDuper.vbs         # Main script (~2000 lines)
├── README.md           # Project overview and usage guide
├── CHANGELOG.md        # Version history (Keep a Changelog format)
├── CONTRIBUTING.md     # This file
├── DEVELOPMENT_PROGRESS.md  # Status and roadmap
├── INSTALLATION.md     # Installation instructions
├── LICENSE             # GPL v3 license text
└── .gitignore          # Git ignore rules
```

## Coding Standards

### General

- **Indentation:** 4 spaces per level. No tabs.
- **Line length:** Keep lines under 120 characters where practical.
- **Option Explicit:** Must be present at the top of the script (`Option Explicit` at line 99).
- **Comments:** Use single-quote (`'`) for all comments. Explain *why*, not *what*.
- **End-of-line style:** CRLF (`\r\n`) — Windows standard.

### Variable Declaration

All variables **must** be declared with `Dim` at the top of their scope. Group related variables and comment their purpose:

```vbscript
Dim i           ' Loop counter
Dim tKind       ' Track kind (KIND_FILE or KIND_STREAM)
Dim dupesFound  ' Count of duplicates detected in current pass
```

### Control Flow

```vbscript
' If/Then/Else
If condition Then
    DoSomething()
Else
    DoOther()
End If

' For loops
For i = 0 To collection.Count - 1
    Process(collection(i))
Next

' While/Wend (prefer Do While/Loop)
Do While Not rs.EOF
    Process(rs.Fields("Name"))
    rs.MoveNext
Loop
```

## Naming Conventions

| Category         | Convention                      | Example                  |
|------------------|---------------------------------|--------------------------|
| Constants        | `UPPER_SNAKE_CASE`              | `DUPE_LOGICAL`           |
| Global variables | `PascalCase`                    | `Intro`, `Outro`         |
| Local variables  | `camelCase`                     | `tKind`, `dupesFound`    |
| Functions/Subs   | `PascalCase`                    | `DedupeTracks`, `Recycle`|
| Objects          | `PascalCase`                    | `iTunesApp`, `FSO`       |
| Dictionaries     | `PascalCase`                    | `DeathRow`, `PlayDupes`  |

## Constants and Magic Values

The script defines named constants at module level (around line 128) to replace magic numbers. **Always use these constants** in new code:

### Duplicate Types (`DUPE_*`)

| Constant              | Value | Description                                    |
|-----------------------|-------|------------------------------------------------|
| `DUPE_LOGICAL`        | 1     | Same file path appears more than once          |
| `DUPE_PHYSICAL`       | 2     | Same metadata + size, different physical files |
| `DUPE_ALTERNATE`      | 3     | Same metadata, different file size             |
| `DUPE_MISSING`        | 4     | Same metadata but file is missing              |
| `DUPE_SAME_SONG`      | 5     | Same artist + name, different album            |
| `DUPE_PSEUDO_PHYS`    | 6     | Physical dupe detected via alternate path      |

### Scan Passes (`PASS_*`)

| Constant              | Value | Pass Description                               |
|-----------------------|-------|-------------------------------------------------|
| `PASS_PLAYLIST`       | 0     | Playlist-level duplicates (same PersistentID)   |
| `PASS_LOGICAL`        | 1     | Logical duplicates (same file path)             |
| `PASS_PHYSICAL`       | 2     | Physical duplicates (same metadata + size)      |
| `PASS_ALTERNATE`      | 3     | Alternate duplicates (same metadata, diff size) |
| `PASS_MISSING`        | 4     | Missing duplicates (metadata match, no file)    |
| `PASS_SAME_SONG`      | 5     | Same song duplicates (artist + name match)      |

### Track Kinds (`KIND_*`)

| Constant        | Value | Description                    |
|-----------------|-------|--------------------------------|
| `KIND_FILE`     | 1     | Track maps to a local file     |
| `KIND_STREAM`   | 3     | Track is a network stream      |

**Do not** use bare literal integers where a named constant exists.

## Error Handling

The script uses `On Error Resume Next` for COM interaction and file operations. Follow these patterns:

```vbscript
' Guarded operation
On Error Resume Next
    result = SomeRiskyOperation()
On Error Goto 0

' Check result and build error message outside the guard
If Err.Number <> 0 Then
    errorMessage = "Failed: " & Err.Description
    Err.Clear
    ' Handle error...
End If
```

- Always pair `On Error Resume Next` with a narrow scope.
- Do not perform string formatting or `MsgBox` calls inside an error guard block.
- Use a flag variable to signal errors out of guarded blocks.

## Testing

### Manual Testing (Windows)

1. Back up your iTunes library before testing.
2. Use **Test Mode** (`Testing=True`) to preview changes without actual deletion.
3. Use **Uncheck mode** (`Uncheck=True`) to mark potential discards for review.
4. Test with a small selection of tracks before running on the full library.

### Code Review Checklist

- [ ] Does the change use named constants instead of magic numbers?
- [ ] Are all new variables declared with `Dim`?
- [ ] Are `On Error Resume Next` blocks kept as narrow as possible?
- [ ] Are error messages built outside the error guard scope?
- [ ] Is the `RegExp` engine reused rather than recreated per iteration?
- [ ] Are COM property accesses cached (e.g., `tKind = T.Kind` before loop body)?
- [ ] Do comments explain *why* the change is needed?

## Pull Request Process

1. Ensure your branch is up to date with `main`:
   ```bash
   git fetch origin
   git rebase origin/main
   ```
2. Run through the [Code Review Checklist](#code-review-checklist).
3. **One change per PR.** If you have multiple independent changes, open separate PRs.
4. Update `CHANGELOG.md` if the change is user-facing (new feature, bug fix, breaking change).
5. Update `README.md` if usage instructions or configuration options change.
6. Open the PR with a clear title and description explaining:
   - What problem the change solves
   - How it was tested
   - Any known limitations

## Commit Conventions

Use [Conventional Commits](https://www.conventionalcommits.org/) prefixes:

| Prefix       | Usage                                               |
|-------------|------------------------------------------------------|
| `feat:`      | A new feature or capability                          |
| `fix:`       | A bug fix                                            |
| `refactor:`  | Code restructuring without functional change         |
| `perf:`      | Performance improvement                              |
| `docs:`      | Documentation-only changes                           |
| `chore:`     | Build, tooling, or dependency changes                |
| `test:`      | Adding or updating tests                             |

Example:
```
feat: add archive-to-folder option for deleted files
fix: handle network path edge case in GetMediaPath
docs: update README with new configuration option
```

---

Thank you for contributing! 🎵
