#!/usr/bin/env python3
"""
Disk Cleanup Script for Windows
Clears cache, temp files, npm cache, git cache, and identifies large folders
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path
from datetime import datetime, timedelta
import json

class DiskCleanup:
    def __init__(self):
        self.stats = {
            'deleted_size': 0,
            'files_deleted': 0,
            'errors': []
        }
        self.dry_run = True
        self.username = os.getenv('USERNAME')
        self.users_path = Path('C:\\Users')
        
    def get_size(self, path):
        """Calculate total size of a directory in MB"""
        try:
            total = sum(f.stat().st_size for f in Path(path).rglob('*') if f.is_file())
            return total / (1024 * 1024)  # Convert to MB
        except:
            return 0
    
    def safe_delete(self, path, description=""):
        """Safely delete files/directories"""
        try:
            path = Path(path)
            if not path.exists():
                return
            
            size_mb = self.get_size(path) if path.is_dir() else path.stat().st_size / (1024 * 1024)
            
            if self.dry_run:
                print(f"  [DRY RUN] Would delete: {path} ({size_mb:.2f} MB) {description}")
                self.stats['deleted_size'] += size_mb
                return
            
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
                print(f"  ✓ Deleted: {path} ({size_mb:.2f} MB)")
            else:
                path.unlink()
                print(f"  ✓ Deleted: {path} ({size_mb:.2f} MB)")
            
            self.stats['deleted_size'] += size_mb
            self.stats['files_deleted'] += 1
        except Exception as e:
            self.stats['errors'].append(f"{path}: {str(e)}")
            print(f"  ✗ Error deleting {path}: {str(e)}")
    
    def clean_windows_temp(self):
        """Clean Windows temporary files"""
        print("\n[1] Cleaning Windows Temp Files...")
        temp_paths = [
            os.path.expandvars('%TEMP%'),
            os.path.expandvars('%TMP%'),
            'C:\\Windows\\Temp',
            f'C:\\Users\\{self.username}\\AppData\\Local\\Temp',
        ]
        
        for temp_path in temp_paths:
            if os.path.exists(temp_path):
                print(f"  Scanning: {temp_path}")
                try:
                    for item in Path(temp_path).iterdir():
                        try:
                            if item.is_file():
                                self.safe_delete(item)
                            elif item.is_dir():
                                self.safe_delete(item)
                        except:
                            pass
                except:
                    pass
    
    def clean_browser_cache(self):
        """Clean browser caches"""
        print("\n[2] Cleaning Browser Caches...")
        
        cache_paths = {
            'Chrome': f'C:\\Users\\{self.username}\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache',
            'Chrome Local': f'C:\\Users\\{self.username}\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\Code Cache',
            'Edge': f'C:\\Users\\{self.username}\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Cache',
            'Firefox': f'C:\\Users\\{self.username}\\AppData\\Local\\Mozilla\\Firefox\\Profiles',
        }
        
        for browser, cache_path in cache_paths.items():
            if os.path.exists(cache_path):
                print(f"  Cleaning {browser} cache...")
                self.safe_delete(cache_path, f"({browser} cache)")
    
    def clean_npm_cache(self):
        """Clean npm cache"""
        print("\n[3] Cleaning NPM Cache...")
        try:
            if shutil.which('npm'):
                if self.dry_run:
                    print("  [DRY RUN] Would run: npm cache clean --force")
                    npm_cache = Path(os.path.expandvars('%APPDATA%\\npm-cache'))
                    if npm_cache.exists():
                        size = self.get_size(npm_cache)
                        print(f"  [DRY RUN] NPM cache size: {size:.2f} MB")
                        self.stats['deleted_size'] += size
                else:
                    result = subprocess.run(['npm', 'cache', 'clean', '--force'], 
                                          capture_output=True, text=True, timeout=60)
                    if result.returncode == 0:
                        print("  ✓ NPM cache cleaned")
                    else:
                        print(f"  ✗ NPM cache clean failed: {result.stderr}")
            else:
                print("  ⚠ npm not found in PATH")
        except Exception as e:
            print(f"  ✗ Error cleaning npm cache: {str(e)}")
    
    def clean_git_cache(self):
        """Find and clean .git directories in projects"""
        print("\n[4] Finding Large .git Directories...")
        
        search_paths = [
            f'C:\\Users\\{self.username}\\Desktop',
            f'C:\\Users\\{self.username}\\Documents',
            f'C:\\Users\\{self.username}\\Downloads',
            f'C:\\Users\\{self.username}',
        ]
        
        git_dirs = []
        for search_path in search_paths:
            if os.path.exists(search_path):
                try:
                    for git_dir in Path(search_path).rglob('.git'):
                        if git_dir.is_dir():
                            size = self.get_size(git_dir)
                            if size > 10:  # Only show .git dirs > 10MB
                                git_dirs.append((git_dir, size))
                except:
                    pass
        
        if git_dirs:
            git_dirs.sort(key=lambda x: x[1], reverse=True)
            print(f"  Found {len(git_dirs)} large .git directories:")
            for git_dir, size in git_dirs[:10]:
                print(f"    - {git_dir.parent}: {size:.2f} MB")
            
            if not self.dry_run:
                response = input("\n  Delete these .git directories? (y/n): ").lower()
                if response == 'y':
                    for git_dir, _ in git_dirs:
                        self.safe_delete(git_dir, "(git repo cache)")
        else:
            print("  ℹ No large .git directories found")
    
    def clean_old_downloads(self):
        """Clean files older than 90 days in Downloads"""
        print("\n[5] Cleaning Old Downloads (>90 days old)...")
        
        downloads_path = Path(f'C:\\Users\\{self.username}\\Downloads')
        if not downloads_path.exists():
            print("  ⚠ Downloads folder not found")
            return
        
        cutoff_date = datetime.now() - timedelta(days=90)
        old_files = []
        
        for item in downloads_path.iterdir():
            try:
                mod_time = datetime.fromtimestamp(item.stat().st_mtime)
                if mod_time < cutoff_date:
                    size = item.stat().st_size / (1024 * 1024) if item.is_file() else self.get_size(item)
                    old_files.append((item, size, mod_time))
            except:
                pass
        
        if old_files:
            old_files.sort(key=lambda x: x[2])
            print(f"  Found {len(old_files)} files/folders older than 90 days:")
            for item, size, mod_time in old_files[:10]:
                print(f"    - {item.name}: {size:.2f} MB (modified: {mod_time.date()})")
            
            if not self.dry_run:
                response = input("\n  Delete these old files? (y/n): ").lower()
                if response == 'y':
                    for item, _, _ in old_files:
                        self.safe_delete(item, "(old download)")
        else:
            print("  ℹ No old files found")
    
    def analyze_user_profiles(self):
        """Analyze space used by each user profile"""
        print("\n[6] Analyzing User Profiles...")
        
        profile_sizes = {}
        for user_dir in self.users_path.iterdir():
            if user_dir.is_dir() and user_dir.name not in ['Default', 'Default User', 'All Users', 'Public']:
                try:
                    size = self.get_size(user_dir)
                    profile_sizes[user_dir.name] = size
                except:
                    pass
        
        if profile_sizes:
            sorted_profiles = sorted(profile_sizes.items(), key=lambda x: x[1], reverse=True)
            print(f"  User profiles by size:")
            for user, size in sorted_profiles:
                print(f"    - {user}: {size:.2f} MB")
        else:
            print("  ⚠ No user profiles found")
    
    def clean_appdata_cache(self):
        """Clean AppData cache directories"""
        print("\n[7] Cleaning AppData Cache...")
        
        appdata_path = Path(f'C:\\Users\\{self.username}\\AppData\\Local')
        cache_patterns = ['Cache', 'cache', 'Temp', 'temp', 'CrashDumps', 'pip', 'pip-cache']
        
        for pattern in cache_patterns:
            for cache_dir in appdata_path.rglob(pattern):
                if cache_dir.is_dir() and cache_dir.parent.name in ['Local', 'LocalLow']:
                    size = self.get_size(cache_dir)
                    if size > 1:  # Only clean if > 1 MB
                        self.safe_delete(cache_dir, f"({pattern} cache)")
    
    def show_summary(self):
        """Show cleanup summary"""
        print("\n" + "="*60)
        print("CLEANUP SUMMARY")
        print("="*60)
        print(f"Mode: {'DRY RUN (no changes made)' if self.dry_run else 'LIVE RUN'}")
        print(f"Total Space Freed: {self.stats['deleted_size']:.2f} MB ({self.stats['deleted_size']/1024:.2f} GB)")
        print(f"Files/Folders Deleted: {self.stats['files_deleted']}")
        
        if self.stats['errors']:
            print(f"\nErrors ({len(self.stats['errors'])}):")
            for error in self.stats['errors'][:5]:
                print(f"  - {error}")
    
    def run(self, dry_run=True, interactive=True):
        """Run all cleanup tasks"""
        self.dry_run = dry_run
        
        print("="*60)
        print("WINDOWS DISK CLEANUP SCRIPT")
        print("="*60)
        print(f"Mode: {'DRY RUN' if dry_run else 'LIVE'}")
        print(f"Current User: {self.username}")
        
        # Show current disk space
        import ctypes
        total, used, free = shutil.disk_usage('C:\\')
        print(f"\nDisk Space (C:)")
        print(f"  Total: {total/(1024**3):.2f} GB")
        print(f"  Used: {used/(1024**3):.2f} GB")
        print(f"  Free: {free/(1024**3):.2f} GB")
        
        if interactive and dry_run:
            response = input("\nRunning in DRY RUN mode. Review changes and confirm. Continue? (y/n): ")
            if response.lower() != 'y':
                print("Aborted.")
                return
        
        # Run cleanup tasks
        self.clean_windows_temp()
        self.clean_browser_cache()
        self.clean_npm_cache()
        self.clean_appdata_cache()
        self.clean_git_cache()
        self.clean_old_downloads()
        self.analyze_user_profiles()
        
        self.show_summary()
        
        # Show disk space after (for dry run)
        if dry_run:
            print(f"\nEstimated new free space: {(free + self.stats['deleted_size'] * 1024 * 1024) / (1024**3):.2f} GB")


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Windows Disk Cleanup Tool')
    parser.add_argument('--live', action='store_true', help='Run in LIVE mode (actually delete files)')
    parser.add_argument('--no-interact', action='store_true', help='Run without interactive prompts')
    parser.add_argument('--json', action='store_true', help='Output results as JSON')
    
    args = parser.parse_args()
    
    cleanup = DiskCleanup()
    cleanup.run(dry_run=not args.live, interactive=not args.no_interact)
    
    if args.json:
        print("\nJSON Output:")
        print(json.dumps(cleanup.stats, indent=2))


if __name__ == '__main__':
    main()
