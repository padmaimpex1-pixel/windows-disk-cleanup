# Windows Disk Cleanup Utility

A comprehensive Python script for Windows system maintenance and disk cleanup.

## Features

✅ **Safe dry-run mode by default** - Preview changes before deleting
✅ **Clears Windows temp files** - %TEMP%, %TMP%, C:\Windows\Temp  
✅ **Clears browser caches** - Chrome, Edge, Firefox  
✅ **Cleans npm cache** - `npm cache clean --force`
✅ **Cleans AppData cache** - Temp, Cache, CrashDumps, pip cache  
✅ **Analyzes disk usage** - Shows size of all user profiles
✅ **Finds large .git directories** - Identifies git repo caches
✅ **Cleans old downloads** - Removes files older than 90 days
✅ **Interactive confirmation** - Review changes before execution
✅ **Detailed reporting** - Shows space freed and errors
✅ **JSON output** - Machine-readable results

## Requirements

- Python 3.6+
- Windows OS
- Administrator privileges (recommended)

## Installation

```bash
git clone https://github.com/padmaimpex1-pixel/windows-disk-cleanup.git
cd windows-disk-cleanup
```

## Usage

### Dry-Run Mode (Preview Only - Recommended First Step)
```bash
python cleanup_disk.py
```

### Live Mode (Actually Delete Files)
```bash
python cleanup_disk.py --live
```

### Non-Interactive Mode
```bash
python cleanup_disk.py --live --no-interact
```

### JSON Output
```bash
python cleanup_disk.py --json
```

### Combine Options
```bash
python cleanup_disk.py --live --json
```

## What Gets Cleaned

| Category | Locations |
|----------|-----------|
| **Temp Files** | %TEMP%, %TMP%, C:\Windows\Temp |
| **Browser Cache** | Chrome, Edge, Firefox default caches |
| **NPM Cache** | npm-cache directory |
| **AppData Cache** | Local cache, temp, CrashDumps, pip |
| **Git Repos** | .git directories > 10MB |
| **Old Downloads** | Files > 90 days old in Downloads |

## Safety Features

- **Dry-run by default** - Run without `--live` to preview changes
- **Error handling** - Safely handles permission errors and missing files
- **User confirmation** - Interactive prompts before critical deletions
- **Size reporting** - Shows exactly what will be deleted and how much space freed
- **Backup-friendly** - Won't delete files needed for browser passwords/bookmarks

## Typical Results

After running this script, you can expect to free:

- **Browser cache**: 100-500 MB (safe to delete)
- **NPM cache**: 50-200 MB (will be rebuilt on demand)
- **Windows temp**: 10-100 MB
- **AppData cache**: 50-300 MB
- **Old downloads**: 100-1000+ MB (depends on your Downloads folder)

**Total: 300-2000+ MB depending on your system**

## Performance Tips

1. **Run regularly** (monthly) to keep disk usage low
2. **Always do dry-run first** to see what will be deleted
3. **Use `--live` only when confident** in the changes
4. **Run as Administrator** for best results
5. **Consider weekly cleanups** if you develop software (npm/git heavy)

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Permission denied" | Run as Administrator |
| Files not deleted | Close applications using those directories |
| npm command not found | Ensure Node.js/npm is installed and in PATH |
| Script very slow | This is normal - analyzing large directories takes time |

## Example Output

```
============================================================
WINDOWS DISK CLEANUP SCRIPT
============================================================
Mode: DRY RUN
Current User: dell

Disk Space (C:)
  Total: 140.00 GB
  Used: 128.66 GB
  Free: 1.59 GB

[1] Cleaning Windows Temp Files...
  Scanning: C:\Users\dell\AppData\Local\Temp
  [DRY RUN] Would delete: C:\Users\dell\AppData\Local\Temp\file.tmp (2.34 MB)

[2] Cleaning Browser Caches...
  Cleaning Chrome cache...
  [DRY RUN] Would delete: C:\Users\dell\AppData\Local\Google\Chrome... (245.67 MB)

============================================================
CLEANUP SUMMARY
============================================================
Mode: DRY RUN (no changes made)
Total Space Freed: 1250.45 MB (1.22 GB)
Files/Folders Deleted: 0
```

## License

MIT License - Feel free to use, modify, and distribute

## Contributing

Contributions welcome! Feel free to:
- Report bugs
- Suggest improvements
- Submit pull requests
- Add platform-specific improvements

## Disclaimer

This script modifies your system by deleting files. While it includes safety measures, **always run in dry-run mode first** and review the changes before using `--live` mode.
