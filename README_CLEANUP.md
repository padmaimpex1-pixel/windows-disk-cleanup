# Disk Cleanup & Utilities

A collection of utility scripts for Windows system maintenance and disk cleanup.

## Scripts

### cleanup_disk.py
Comprehensive Windows disk cleanup tool that:
- Clears Windows temp files
- Clears browser caches (Chrome, Edge, Firefox)
- Cleans npm cache
- Cleans AppData cache
- Identifies large .git directories
- Cleans old downloads (>90 days)
- Analyzes user profiles

**Usage:**

```bash
# Dry-run mode (preview changes, no deletion)
python cleanup_disk.py

# Live mode (actually delete files)
python cleanup_disk.py --live

# Non-interactive
python cleanup_disk.py --live --no-interact

# JSON output
python cleanup_disk.py --json
```

## Requirements
- Python 3.6+
- Windows OS
- Administrator privileges recommended

## Features
✅ Safe dry-run mode by default
✅ Interactive confirmation prompts
✅ Detailed reporting of freed space
✅ Error handling and recovery
✅ Supports multiple user profiles
✅ JSON output support

## Installation

Clone this repository:
```bash
git clone https://github.com/padmaimpex1-pixel/starter-repo.git
cd starter-repo
```

Run the cleanup script:
```bash
python cleanup_disk.py
```

## Safety Notes
- **Always run in dry-run mode first** to preview changes
- The script uses error handling to prevent accidental deletion of important files
- Browser caches can be safely deleted without affecting saved passwords
- Old files are filtered by modification date to prevent data loss

## License
MIT License
