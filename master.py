"""
master.py — padmaimpex1-pixel workspace master script.
Each task is a function. Run a specific function or run all.
Output goes to D:\\Generated-Outputs by default.
"""

import os
import shutil
import sqlite3
import tempfile
import time
import json
from datetime import datetime
from pathlib import Path

OUTPUT_DIR = Path(r"D:\Generated-Outputs")
COPILOT_DB = Path(r"C:\Users\dell\AppData\Roaming\Code\User\globalStorage\github.copilot-chat\session-store.db")


# ---------------------------------------------------------------------------
# TASK: Export all Copilot session prompts + responses to Excel
# ---------------------------------------------------------------------------

def export_copilot_sessions_to_excel(output_path: Path = None, watch: bool = False, interval_seconds: int = 60):
    """
    Reads all Copilot chat sessions and turns from the local session-store SQLite DB
    and writes them to an Excel file with multiple sheets categorized by prompt type.

    Args:
        output_path: Where to save the .xlsx file. Defaults to OUTPUT_DIR/copilot_sessions.xlsx
        watch:       If True, re-exports continuously every `interval_seconds`.
        interval_seconds: How often to re-export when watch=True.
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("openpyxl not installed. Run: pip install openpyxl")
        return

    if output_path is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / "copilot_sessions.xlsx"

    output_path = Path(output_path)

    print(f"Copilot DB : {COPILOT_DB}")
    print(f"Output XLS : {output_path}")

    if not COPILOT_DB.exists():
        print(f"ERROR: session-store.db not found at {COPILOT_DB}")
        return

    def categorize_prompt(prompt_text):
        """Categorize prompt by keywords."""
        if not prompt_text:
            return "Other"
        p = prompt_text.lower()
        
        # Social Media
        if any(x in p for x in ['twitter', 'facebook', 'instagram', 'tiktok', 'reddit', 'linkedin', 
                                'discord', 'slack', 'whatsapp', 'telegram', 'snapchat', 'pinterest',
                                'youtube', 'twitch', 'mastodon', 'bluesky', 'threads', 'x.com',
                                'social', 'post', 'tweet', 'share', 'viral', 'engagement', 'followers',
                                'hashtag', '@mention', 'dm ', 'dm:', 'channel', 'community', 'group chat']):
            return "Social Media"
        
        # Videos
        if any(x in p for x in ['video', 'mp4', 'avi', 'mov', 'mkv', 'webm', 'flv', 'wmv',
                                'codec', 'ffmpeg', 'premiere', 'davinci', 'resolve', 'final cut',
                                'after effects', 'camtasia', 'obs', 'streaming', 'youtube', 'vimeo',
                                'encoding', 'transcoding', 'resolution', '1080p', '4k', '8k',
                                'frame rate', 'fps', 'bitrate', 'h.264', 'h.265', 'vp9',
                                'video editing', 'motion', 'animation', 'subtitle', 'caption']):
            return "Videos"
        
        # Photos & Images
        if any(x in p for x in ['photo', 'image', 'picture', 'jpg', 'jpeg', 'png', 'gif', 'bmp',
                                'svg', 'webp', 'tiff', 'raw', 'psd', 'ai', 'eps', 'pdf',
                                'photoshop', 'gimp', 'lightroom', 'capture one', 'affinity',
                                'exposure', 'aperture', 'shutter', 'iso', 'white balance',
                                'pixel', 'resolution', 'dpi', 'crop', 'rotate', 'filter',
                                'photography', 'photographic', 'screenshot', 'screenshot', 'canvas',
                                'design', 'graphic', 'visual', 'illustration', 'artwork']):
            return "Photos & Images"
        
        # URLs (check after media categories)
        if any(proto in p for proto in ['http://', 'https://', 'ftp://', 'ssh://', '.com', '.org', '.net', 'github.com', 'gitlab.com']):
            if any(x in p for x in ['http', 'https', 'ftp', '.com', '.org', '.net', 'url', 'link', 'web']):
                return "URLs & Links"
        
        # Ports
        if any(x in p for x in ['port', 'localhost:', ':8080', ':3000', ':5000', ':443', ':80', 'listen', 'binding', 'socket']):
            return "Ports & Networking"
        
        # Backup operations
        if any(x in p for x in ['backup', 'restore', 'archive', 'compress', 'zip', 'tar', 'dump', 'snapshot']):
            return "Backup & Archive"
        
        # Linux/Bash commands (check before generic script detection)
        linux_cmds = [
            'ls ', 'cat ', 'grep ', 'sed ', 'awk ', 'chmod ', 'chown ', 'mkdir ', 'rmdir ',
            'find ', 'locate ', 'which ', 'whereis ', 'apt ', 'yum ', 'pacman ', 'dnf ',
            'systemctl', 'service ', 'sudo ', 'su ', 'ssh ', 'scp ', 'sftp', 'rsync',
            'tar ', 'gzip ', 'bzip2 ', 'unzip ', 'gunzip ', 'tail ', 'head ', 'wc ',
            'cut ', 'paste ', 'sort ', 'uniq ', 'diff ', 'patch ', 'cmp ', 'comm ',
            'file ', 'strings ', 'ldd ', 'nm ', 'readelf ', 'objdump ', 'strace ',
            'ltrace ', 'top ', 'htop ', 'ps ', 'kill ', 'killall ', 'pkill ', 'pgrep ',
            'df ', 'du ', 'mount ', 'umount ', 'fsck ', 'mkfs ', 'fdisk ', 'parted',
            'ifconfig', 'ip addr', 'route ', 'iptables', 'firewall', 'netcat', 'nc ',
            '/bin/', '/usr/bin/', '/etc/', '/var/', '/home/', 'bash', 'sh ', 'zsh'
        ]
        if any(cmd in p for cmd in linux_cmds):
            return "Linux Commands"
        
        # PowerShell commands (check before generic script detection)
        powershell_cmds = [
            'get-childitem', 'get-content', 'select-string', 'invoke-webrequest', 
            'start-process', 'stop-process', 'new-item', 'remove-item', 'copy-item',
            'move-item', 'rename-item', 'set-location', 'test-path', 'get-process',
            'write-host', 'read-host', 'foreach-object', 'where-object', 'group-object',
            'sort-object', 'measure-object', 'get-service', 'start-service', 'stop-service',
            'get-eventlog', 'get-wmiobject', 'add-content', 'clear-content', 'import-module',
            'export-csv', 'import-csv', 'convertto-json', 'convertfrom-json', 'new-psdrive'
        ]
        if any(cmd in p for cmd in powershell_cmds):
            return "PowerShell Commands"
        
        # Windows commands (check before generic script detection)
        windows_cmds = [
            'cmd ', 'dir ', 'del ', 'copy ', 'move ', 'rename ', 'ren ', 'type ', 
            'tasklist', 'taskkill', 'systeminfo', 'ipconfig', 'ping ', 'netstat', 
            'schtasks', 'wmic', 'reg ', 'bitsadmin', 'findstr', 'fc ', 'for /f',
            'pushd', 'popd', 'cd ', 'mkdir', 'rmdir', 'robocopy', 'xcopy', 'comp ',
            'diskpart', 'format ', 'chkdsk', 'defrag', 'cipher', 'takeown', 'icacls',
            'attrib ', 'date /t', 'time /t', 'nslookup', 'tracert', 'route ', 'arp '
        ]
        if any(cmd in p for cmd in windows_cmds):
            return "Windows Commands"
        
        # Repository-specific work (check for common repo patterns)
        repo_patterns = [
            'starter-repo', 'windows-disk-cleanup', 'pixel', 'padmaimpex', 'myapp',
            'vipingit', 'misc', 'd:\\gitrepos', 'github.com/', 'gitlab.com/', 'bitbucket.org/',
            'src/', 'test/', 'docs/', 'readme.md', '.github/', '.gitignore'
        ]
        if any(repo in p for repo in repo_patterns):
            return "Repository Work"
        
        if any(x in p for x in ['move', 'copy', 'delete', 'rename', 'organize', 'folder', 'directory']):
            return "File Operations"
        if any(x in p for x in ['python', 'script', 'code', 'function', 'class', 'def']):
            return "Code & Scripts"
        if any(x in p for x in ['git', 'github', 'commit', 'push', 'merge', 'branch']):
            return "Git & GitHub"
        if any(x in p for x in ['install', 'pip', 'npm', 'package', 'dependency', 'requirements']):
            return "Package Management"
        if any(x in p for x in ['fix', 'bug', 'error', 'debug', 'issue', 'problem']):
            return "Debugging & Fixes"
        if any(x in p for x in ['create', 'generate', 'write', 'build', 'implement']):
            return "Creation & Generation"
        if any(x in p for x in ['read', 'view', 'check', 'list', 'show', 'display', 'inspect']):
            return "Inspection & Analysis"
        if any(x in p for x in ['excel', 'xls', 'csv', 'export', 'data', 'database']):
            return "Data & Export"
        if any(x in p for x in ['instruction', 'config', 'settings', 'default', 'convention']):
            return "Configuration & Conventions"
        return "Other"

    def do_export():
        # Copy DB to temp location to avoid locking the live file
        tmp = Path(tempfile.gettempdir()) / "copilot-session-store-copy.db"
        shutil.copy2(COPILOT_DB, tmp)

        con = sqlite3.connect(tmp)
        con.row_factory = sqlite3.Row
        cur = con.cursor()

        rows = cur.execute("""
            SELECT
                s.id            AS session_id,
                s.cwd           AS cwd,
                s.repository    AS repository,
                s.branch        AS branch,
                s.summary       AS session_summary,
                s.created_at    AS session_created,
                s.updated_at    AS session_updated,
                t.turn_index    AS turn_index,
                t.user_message  AS prompt,
                t.assistant_response AS response,
                t.timestamp     AS turn_timestamp
            FROM sessions s
            JOIN turns t ON t.session_id = s.id
            ORDER BY t.timestamp DESC, s.created_at DESC, t.turn_index DESC
        """).fetchall()

        con.close()
        tmp.unlink(missing_ok=True)

        # Group by category
        categorized = {}
        for row in rows:
            cat = categorize_prompt(row['prompt'])
            if cat not in categorized:
                categorized[cat] = []
            categorized[cat].append(row)

        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        headers = [
            "Session ID", "CWD", "Repository", "Branch",
            "Session Summary", "Session Created", "Session Updated",
            "Turn #", "Prompt (User)", "Response (Copilot)", "Turn Timestamp"
        ]

        # Summary sheet
        ws_summary = wb.create_sheet("Summary", 0)
        ws_summary['A1'] = "Copilot Sessions Report"
        ws_summary['A1'].font = Font(bold=True, size=14)
        ws_summary['A3'] = "Category"
        ws_summary['B3'] = "Turn Count"
        ws_summary['A3'].font = Font(bold=True)
        ws_summary['B3'].font = Font(bold=True)
        row_num = 4
        for cat in sorted(categorized.keys()):
            ws_summary[f'A{row_num}'] = cat
            ws_summary[f'B{row_num}'] = len(categorized[cat])
            row_num += 1
        ws_summary['A2'] = f"Total turns: {len(rows)} | Total categories: {len(categorized)}"
        
        # Keywords sheet
        ws_keywords = wb.create_sheet("Keywords", 1)
        ws_keywords['A1'] = "Keyword Analysis"
        ws_keywords['A1'].font = Font(bold=True, size=14)
        
        keyword_dict = {}
        keyword_patterns = [
            'python', 'javascript', 'typescript', 'java', 'csharp', 'rust', 'go', 'rust',
            'docker', 'kubernetes', 'aws', 'azure', 'gcp', 'cloud', 'database', 'sql', 'nosql',
            'api', 'rest', 'graphql', 'websocket', 'microservice', 'serverless', 'lambda',
            'testing', 'unit test', 'integration test', 'e2e', 'jest', 'pytest', 'mocha',
            'ci/cd', 'jenkins', 'gitlab-ci', 'github actions', 'devops', 'terraform',
            'monitoring', 'logging', 'prometheus', 'grafana', 'elk', 'datadog', 'newrelic',
            'security', 'encryption', 'ssl', 'tls', 'oauth', 'jwt', 'authentication',
            'performance', 'optimization', 'cache', 'redis', 'memcached', 'caching',
            'frontend', 'react', 'vue', 'angular', 'svelte', 'html', 'css', 'webpack',
            'backend', 'nodejs', 'django', 'flask', 'fastapi', 'spring', 'rails',
            'database', 'postgresql', 'mysql', 'mongodb', 'dynamodb', 'cassandra',
            'version control', 'git', 'github', 'gitlab', 'bitbucket', 'svn',
            'linux', 'ubuntu', 'debian', 'centos', 'rhel', 'alpine', 'docker',
            'windows', 'powershell', 'batch', 'cmd', 'wsl', 'hyper-v',
            'mac', 'macos', 'osx', 'homebrew', 'xcode', 'swift', 'objective-c'
        ]
        
        for row in rows:
            prompt = (row['prompt'] or '').lower()
            response = (row['response'] or '').lower()
            combined = prompt + ' ' + response
            for kw in keyword_patterns:
                if kw in combined:
                    if kw not in keyword_dict:
                        keyword_dict[kw] = {'count': 0, 'sessions': set(), 'categories': []}
                    keyword_dict[kw]['count'] += 1
                    keyword_dict[kw]['sessions'].add(row['session_id'])
                    cat = categorize_prompt(row['prompt'])
                    if cat not in keyword_dict[kw]['categories']:
                        keyword_dict[kw]['categories'].append(cat)
        
        # Sort keywords by frequency
        sorted_keywords = sorted(keyword_dict.items(), key=lambda x: x[1]['count'], reverse=True)
        
        ws_keywords['A3'] = "Keyword"
        ws_keywords['B3'] = "Count"
        ws_keywords['C3'] = "Sessions"
        ws_keywords['D3'] = "Categories"
        for col in ['A', 'B', 'C', 'D']:
            ws_keywords[f'{col}3'].font = Font(bold=True)
            ws_keywords[f'{col}3'].fill = PatternFill("solid", fgColor="1F4E79")
            ws_keywords[f'{col}3'].font = Font(bold=True, color="FFFFFF")
        
        row_num = 4
        for kw, data in sorted_keywords:
            ws_keywords[f'A{row_num}'] = kw
            ws_keywords[f'B{row_num}'] = data['count']
            ws_keywords[f'C{row_num}'] = len(data['sessions'])
            ws_keywords[f'D{row_num}'] = ', '.join(data['categories'][:3])
            row_num += 1
        
        ws_keywords.column_dimensions['A'].width = 20
        ws_keywords.column_dimensions['B'].width = 10
        ws_keywords.column_dimensions['C'].width = 12
        ws_keywords.column_dimensions['D'].width = 50

        # One sheet per category
        for cat in sorted(categorized.keys()):
            cat_rows = categorized[cat]
            sheet_name = cat[:31]  # Excel sheet name limit
            ws = wb.create_sheet(sheet_name)

            # Header row styling
            header_fill = PatternFill("solid", fgColor="1F4E79")
            header_font = Font(bold=True, color="FFFFFF", size=10)
            for col, h in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(wrap_text=True, vertical="center")

            # Data rows
            alt_fill = PatternFill("solid", fgColor="D9E1F2")
            for r_idx, row in enumerate(cat_rows, start=2):
                fill = alt_fill if r_idx % 2 == 0 else PatternFill()
                values = [
                    row["session_id"], row["cwd"], row["repository"], row["branch"],
                    row["session_summary"], row["session_created"], row["session_updated"],
                    row["turn_index"], row["prompt"], row["response"], row["turn_timestamp"]
                ]
                for c_idx, val in enumerate(values, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val or "")
                    cell.fill = fill
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

            # Column widths
            col_widths = [36, 30, 40, 12, 40, 22, 22, 8, 60, 80, 22]
            for i, w in enumerate(col_widths, start=1):
                ws.column_dimensions[get_column_letter(i)].width = min(w, 80)

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

        wb.save(output_path)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{now}] Exported {len(rows)} turns across {len(set(r['session_id'] for r in rows))} sessions, {len(categorized)} categories -> {output_path}")

    if watch:
        print(f"Watch mode ON -- refreshing every {interval_seconds}s. Press Ctrl+C to stop.")
        while True:
            try:
                do_export()
                time.sleep(interval_seconds)
            except KeyboardInterrupt:
                print("Stopped.")
                break
    else:
        do_export()


# ---------------------------------------------------------------------------
# TASK: Sync vipingit1 repos to padmaimpex1-pixel on GitHub
# ---------------------------------------------------------------------------

def sync_gitrepos_to_padmaimpex(root: str = r"D:\GitRepos", target_owner: str = "padmaimpex1-pixel", dry_run: bool = False):
    """
    Scans all Git repos under `root`, and for any repo whose remote origin
    points to `vipingit1`, re-points it to `padmaimpex1-pixel` and pushes.

    Skips:
      - Duplicate folders containing '-from-' in their name
      - Third-party repos (not vipingit1 or padmaimpex1-pixel)
      - Repos already on padmaimpex1-pixel (just pushes them)

    Args:
        root:         Root folder to scan for git repos.
        target_owner: GitHub account to push to.
        dry_run:      If True, print what would happen without making changes.
    """
    import subprocess, json, urllib.request, urllib.error

    def run(cmd, cwd=None, capture=True):
        r = subprocess.run(cmd, cwd=cwd, capture_output=capture, text=True)
        return r.stdout.strip(), r.stderr.strip(), r.returncode

    def get_git_token():
        out, _, _ = run(['git', 'credential', 'fill'],)
        # Try to get token from Windows credential manager via git
        # Use a known working URL to extract token
        import subprocess as sp
        proc = sp.Popen(['git', 'credential', 'fill'], stdin=sp.PIPE, stdout=sp.PIPE, stderr=sp.PIPE, text=True)
        stdout, _ = proc.communicate(input="protocol=https\nhost=github.com\n\n")
        for line in stdout.splitlines():
            if line.startswith('password='):
                return line[len('password='):]
        return None

    def repo_exists_on_github(owner, repo_name, token):
        url = f"https://api.github.com/repos/{owner}/{repo_name}"
        req = urllib.request.Request(url, headers={"Authorization": f"token {token}", "User-Agent": "master.py"})
        try:
            urllib.request.urlopen(req, timeout=10)
            return True
        except urllib.error.HTTPError as e:
            return e.code != 404

    def create_repo_on_github(owner, repo_name, token, private=True):
        url = "https://api.github.com/user/repos"
        data = json.dumps({"name": repo_name, "private": private, "auto_init": False}).encode()
        req = urllib.request.Request(url, data=data, method="POST",
                                     headers={"Authorization": f"token {token}",
                                              "Content-Type": "application/json",
                                              "User-Agent": "master.py"})
        try:
            urllib.request.urlopen(req, timeout=15)
            return True, "created"
        except urllib.error.HTTPError as e:
            body = e.read().decode()
            if "already exists" in body or e.code == 422:
                return True, "already_exists"
            return False, f"HTTP {e.code}: {body[:200]}"

    # Find all git repos
    root_path = Path(root)
    git_dirs = [p.parent for p in root_path.rglob('.git') if p.is_dir()]
    git_dirs = sorted(set(git_dirs))

    token = get_git_token()
    if not token:
        print("ERROR: Could not retrieve GitHub token from git credential manager.")
        return

    results = []
    skipped = []

    for repo_path in git_dirs:
        name = repo_path.name

        # Skip duplicates
        if '-from-' in name:
            skipped.append((str(repo_path), "duplicate"))
            continue

        remote_out, _, rc = run(['git', 'remote', 'get-url', 'origin'], cwd=str(repo_path))
        if rc != 0 or not remote_out:
            skipped.append((str(repo_path), "no_remote"))
            continue

        remote = remote_out.strip().rstrip('.git')
        remote_lower = remote.lower()

        # Already on target
        if f"github.com/{target_owner.lower()}/" in remote_lower:
            repo_name = remote.split('/')[-1]
            status = "already_target"
        elif "github.com/vipingit1/" in remote_lower:
            repo_name = remote.split('/')[-1]
            status = "vipingit1"
        else:
            skipped.append((str(repo_path), f"third_party: {remote}"))
            continue

        branch_out, _, _ = run(['git', 'branch', '--show-current'], cwd=str(repo_path))
        branch = branch_out.strip() or "main"
        new_remote = f"https://github.com/{target_owner}/{repo_name}.git"

        print(f"\n[{name}]")
        print(f"  Repo   : {repo_name}")
        print(f"  Remote : {remote} -> {new_remote}")
        print(f"  Branch : {branch}")

        if dry_run:
            print("  DRY_RUN: skipping actual changes")
            results.append({"repo": name, "action": "dry_run", "target": new_remote})
            continue

        # Ensure repo exists on padmaimpex1-pixel
        exists = repo_exists_on_github(target_owner, repo_name, token)
        if not exists:
            ok, msg = create_repo_on_github(target_owner, repo_name, token, private=True)
            print(f"  Create : {msg}")
            if not ok:
                results.append({"repo": name, "action": "create_failed", "error": msg})
                continue
        else:
            print(f"  Create : repo already exists on GitHub")

        # Update remote
        if status == "vipingit1":
            run(['git', 'remote', 'set-url', 'origin', new_remote], cwd=str(repo_path))
            print(f"  Remote updated to {new_remote}")

        # Push
        push_out, push_err, push_rc = run(['git', 'push', 'origin', branch, '--force-with-lease'], cwd=str(repo_path))
        if push_rc == 0:
            print(f"  PUSHED OK")
            results.append({"repo": name, "action": "pushed", "target": new_remote})
        else:
            # Try without --force-with-lease
            push_out2, push_err2, push_rc2 = run(['git', 'push', '--set-upstream', 'origin', branch], cwd=str(repo_path))
            if push_rc2 == 0:
                print(f"  PUSHED OK (set-upstream)")
                results.append({"repo": name, "action": "pushed", "target": new_remote})
            else:
                print(f"  PUSH FAILED: {push_err or push_err2}")
                results.append({"repo": name, "action": "push_failed", "error": push_err or push_err2})

    print(f"\n\n=== SYNC SUMMARY ===")
    pushed    = [r for r in results if r.get('action') == 'pushed']
    failed    = [r for r in results if 'failed' in r.get('action','')]
    print(f"  Pushed  : {len(pushed)}")
    print(f"  Failed  : {len(failed)}")
    print(f"  Skipped : {len(skipped)}")
    if failed:
        print("\n  FAILURES:")
        for f in failed:
            print(f"    {f['repo']}: {f.get('error','')[:120]}")
    if skipped:
        print("\n  SKIPPED:")
        for s, reason in skipped:
            print(f"    {Path(s).name}: {reason}")


# ---------------------------------------------------------------------------
# MENU
# ---------------------------------------------------------------------------

TASKS = {
    "1": ("Export Copilot sessions to Excel (once)", lambda: export_copilot_sessions_to_excel()),
    "2": ("Export Copilot sessions to Excel (watch/continuous)", lambda: export_copilot_sessions_to_excel(watch=True, interval_seconds=60)),
    "3": ("Sync vipingit1 repos to padmaimpex1-pixel (dry run)", lambda: sync_gitrepos_to_padmaimpex(dry_run=True)),
    "4": ("Sync vipingit1 repos to padmaimpex1-pixel (LIVE)", lambda: sync_gitrepos_to_padmaimpex(dry_run=False)),
}

if __name__ == "__main__":
    print("\n=== padmaimpex1-pixel master.py ===\n")
    for key, (label, _) in TASKS.items():
        print(f"  [{key}] {label}")
    print()
    choice = input("Select task (or press Enter for task 1): ").strip() or "1"
    if choice in TASKS:
        TASKS[choice][1]()
    else:
        print(f"Unknown choice: {choice}")


# ---------------------------------------------------------------------------
# TASK: List heavy running user processes
# ---------------------------------------------------------------------------

def list_heavy_processes(output_path: Path = None, threshold_cpu: float = 5.0, threshold_memory: float = 50.0):
    """
    Lists all user processes consuming significant CPU/Memory resources.
    
    Args:
        output_path: Where to save the report. Defaults to OUTPUT_DIR/heavy_processes.xlsx
        threshold_cpu: Show processes using more than this % CPU
        threshold_memory: Show processes using more than this MB memory
    """
    try:
        import psutil
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        print("psutil and openpyxl required. Run: pip install psutil openpyxl")
        return

    if output_path is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / "heavy_processes.xlsx"

    output_path = Path(output_path)
    print(f"Scanning user processes...")

    processes = []
    for proc in psutil.process_iter(['pid', 'name', 'username', 'cpu_percent', 'memory_info']):
        try:
            pinfo = proc.as_dict(attrs=['pid', 'name', 'username', 'cpu_percent', 'memory_info'])
            cpu = pinfo['cpu_percent'] or 0
            mem_mb = (pinfo['memory_info'].rss / (1024 * 1024)) if pinfo['memory_info'] else 0
            
            if cpu > threshold_cpu or mem_mb > threshold_memory:
                processes.append({
                    'PID': pinfo['pid'],
                    'Process Name': pinfo['name'],
                    'User': pinfo['username'] or 'System',
                    'CPU %': round(cpu, 2),
                    'Memory MB': round(mem_mb, 2),
                })
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    # Sort by memory usage
    processes.sort(key=lambda x: x['Memory MB'], reverse=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Heavy Processes"

    headers = ['PID', 'Process Name', 'User', 'CPU %', 'Memory MB']
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font

    # Add summary
    ws.insert_rows(1, 2)
    ws['A1'] = f"Heavy Processes Report - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = f"Threshold: CPU > {threshold_cpu}% or Memory > {threshold_memory}MB | Total: {len(processes)} processes"
    ws['A2'].font = Font(italic=True, size=10)

    alt_fill = PatternFill("solid", fgColor="D9E1F2")
    for r_idx, proc in enumerate(processes, start=4):
        fill = alt_fill if r_idx % 2 == 0 else PatternFill()
        for col, key in enumerate(headers, start=1):
            cell = ws.cell(row=r_idx, column=col, value=proc[key])
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center" if key in ['PID', 'CPU %', 'Memory MB'] else "left")

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 12

    wb.save(output_path)
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Found {len(processes)} heavy processes -> {output_path}")


# ---------------------------------------------------------------------------
# TASK: System monitoring graph/report
# ---------------------------------------------------------------------------

def system_monitoring_report(output_path: Path = None, samples: int = 60, interval: int = 1):
    """
    Captures CPU, Memory, Disk, and Network stats over time and creates a report.
    
    Args:
        output_path: Where to save the report. Defaults to OUTPUT_DIR/system_monitoring.xlsx
        samples: Number of data samples to collect
        interval: Seconds between samples
    """
    try:
        import psutil
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        print("psutil and openpyxl required. Run: pip install psutil openpyxl")
        return

    if output_path is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / "system_monitoring.xlsx"

    output_path = Path(output_path)
    print(f"Collecting system monitoring data ({samples} samples)...")

    data = []
    for i in range(samples):
        timestamp = datetime.now().strftime('%H:%M:%S')
        cpu_percent = psutil.cpu_percent(interval=0.1)
        memory = psutil.virtual_memory()
        disk = psutil.disk_usage('/')
        net = psutil.net_io_counters()
        
        data.append({
            'Timestamp': timestamp,
            'CPU %': cpu_percent,
            'Memory %': memory.percent,
            'Memory Used GB': round(memory.used / (1024**3), 2),
            'Disk %': disk.percent,
            'Disk Used GB': round(disk.used / (1024**3), 2),
            'Net In MB': round(net.bytes_recv / (1024**2), 2),
            'Net Out MB': round(net.bytes_sent / (1024**2), 2),
        })
        
        if i < samples - 1:
            print(f"  [{i+1}/{samples}] CPU: {cpu_percent:.1f}% | Memory: {memory.percent:.1f}% | Disk: {disk.percent:.1f}%")
            time.sleep(interval)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "System Monitoring"

    headers = list(data[0].keys())
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font

    ws['A1'] = "System Monitoring Report"
    ws['A1'].font = Font(bold=True, size=14)

    # Calculate averages
    avg_cpu = sum(d['CPU %'] for d in data) / len(data)
    avg_mem = sum(d['Memory %'] for d in data) / len(data)
    
    ws['A1'].value = f"System Monitoring Report - Avg CPU: {avg_cpu:.1f}%, Avg Mem: {avg_mem:.1f}%"

    alt_fill = PatternFill("solid", fgColor="D9E1F2")
    for r_idx, row_data in enumerate(data, start=3):
        fill = alt_fill if r_idx % 2 == 0 else PatternFill()
        for col, key in enumerate(headers, start=1):
            cell = ws.cell(row=r_idx, column=col, value=row_data[key])
            cell.fill = fill
            cell.alignment = Alignment(horizontal="right")

    col_widths = [12, 10, 10, 15, 10, 15, 12, 12]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[chr(64+i)].width = w

    wb.save(output_path)
    print(f"System monitoring report -> {output_path}")


# ---------------------------------------------------------------------------
# TASK: List all Chrome extensions
# ---------------------------------------------------------------------------

def list_chrome_extensions(output_path: Path = None):
    """
    Lists all installed Chrome/Chromium extensions.
    
    Args:
        output_path: Where to save the report. Defaults to OUTPUT_DIR/chrome_extensions.xlsx
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        print("openpyxl required. Run: pip install openpyxl")
        return

    if output_path is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / "chrome_extensions.xlsx"

    output_path = Path(output_path)

    # Chrome extensions directory
    username = os.getenv('USERNAME')
    chrome_path = Path(rf"C:\Users\{username}\AppData\Local\Google\Chrome\User Data\Default\Extensions")
    
    extensions = []
    
    if chrome_path.exists():
        print(f"Scanning Chrome extensions at {chrome_path}...")
        for ext_dir in chrome_path.iterdir():
            if ext_dir.is_dir():
                try:
                    # Get manifest from latest version folder
                    versions = sorted([d for d in ext_dir.iterdir() if d.is_dir()], reverse=True)
                    if versions:
                        manifest_path = versions[0] / "manifest.json"
                        if manifest_path.exists():
                            with open(manifest_path, 'r', encoding='utf-8') as f:
                                manifest = json.load(f)
                                ext_name = manifest.get('name', 'Unknown')
                                ext_version = manifest.get('version', 'Unknown')
                                ext_desc = manifest.get('description', '')
                                permissions = ', '.join(manifest.get('permissions', [])[:5])
                                
                                extensions.append({
                                    'Extension ID': ext_dir.name,
                                    'Name': ext_name,
                                    'Version': ext_version,
                                    'Description': ext_desc[:100],
                                    'Key Permissions': permissions,
                                })
                except:
                    pass
    else:
        print(f"Chrome path not found at {chrome_path}")
        return

    extensions.sort(key=lambda x: x['Name'])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Chrome Extensions"

    ws['A1'] = f"Chrome Extensions Report - {len(extensions)} extensions"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ['Extension ID', 'Name', 'Version', 'Description', 'Key Permissions']
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font

    alt_fill = PatternFill("solid", fgColor="D9E1F2")
    for r_idx, ext in enumerate(extensions, start=3):
        fill = alt_fill if r_idx % 2 == 0 else PatternFill()
        for col, key in enumerate(headers, start=1):
            cell = ws.cell(row=r_idx, column=col, value=ext[key])
            cell.fill = fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    col_widths = [32, 30, 10, 50, 40]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[chr(64+i)].width = w

    ws.freeze_panes = "A3"
    wb.save(output_path)
    print(f"Found {len(extensions)} Chrome extensions -> {output_path}")


# ---------------------------------------------------------------------------
# TASK: Scan all drives
# ---------------------------------------------------------------------------

def scan_all_drives(output_path: Path = None):
    """
    Scans all drives and reports usage statistics.
    
    Args:
        output_path: Where to save the report. Defaults to OUTPUT_DIR/drives_scan.xlsx
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        import psutil
    except ImportError:
        print("openpyxl and psutil required. Run: pip install openpyxl psutil")
        return

    if output_path is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / "drives_scan.xlsx"

    output_path = Path(output_path)

    drives = []
    print("Scanning all drives...")

    try:
        import psutil
        for partition in psutil.disk_partitions():
            if partition.fstype:  # Skip empty partitions
                try:
                    usage = psutil.disk_usage(partition.mountpoint)
                    used_pct = (usage.used / usage.total * 100) if usage.total > 0 else 0
                    
                    drives.append({
                        'Drive': partition.device,
                        'Mount Point': partition.mountpoint,
                        'File System': partition.fstype,
                        'Total GB': round(usage.total / (1024**3), 2),
                        'Used GB': round(usage.used / (1024**3), 2),
                        'Free GB': round(usage.free / (1024**3), 2),
                        'Used %': round(used_pct, 2),
                    })
                except PermissionError:
                    drives.append({
                        'Drive': partition.device,
                        'Mount Point': partition.mountpoint,
                        'File System': partition.fstype,
                        'Total GB': 'N/A',
                        'Used GB': 'N/A',
                        'Free GB': 'N/A',
                        'Used %': 'Access Denied',
                    })
    except Exception as e:
        print(f"Error scanning drives: {e}")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Drives Scan"

    ws['A1'] = f"Drive Scan Report - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ['Drive', 'Mount Point', 'File System', 'Total GB', 'Used GB', 'Free GB', 'Used %']
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font

    alt_fill = PatternFill("solid", fgColor="D9E1F2")
    warning_fill = PatternFill("solid", fgColor="FFE699")

    for r_idx, drive in enumerate(drives, start=3):
        used_pct = drive['Used %']
        fill = warning_fill if (isinstance(used_pct, (int, float)) and used_pct > 80) else (alt_fill if r_idx % 2 == 0 else PatternFill())
        
        for col, key in enumerate(headers, start=1):
            cell = ws.cell(row=r_idx, column=col, value=drive[key])
            cell.fill = fill
            cell.alignment = Alignment(horizontal="right" if key in ['Total GB', 'Used GB', 'Free GB', 'Used %'] else "left")

    col_widths = [12, 25, 12, 12, 12, 12, 10]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[chr(64+i)].width = w

    wb.save(output_path)
    print(f"Drive scan report ({len(drives)} drives) -> {output_path}")


# ---------------------------------------------------------------------------
# TASK: Scan C drive and list all high size files
# ---------------------------------------------------------------------------

def scan_c_drive_large_files(output_path: Path = None, min_size_mb: float = 100.0, max_results: int = 1000):
    """
    Recursively scans C: drive and lists all large files.
    
    Args:
        output_path: Where to save the report. Defaults to OUTPUT_DIR/large_files_c_drive.xlsx
        min_size_mb: Minimum file size in MB to report (default: 100 MB)
        max_results: Maximum number of files to report (default: 1000)
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
    except ImportError:
        print("openpyxl required. Run: pip install openpyxl")
        return

    if output_path is None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / "large_files_c_drive.xlsx"

    output_path = Path(output_path)

    print(f"Scanning C: drive for files > {min_size_mb}MB...")
    print("This may take several minutes depending on drive size and file count...")

    files = []
    skipped_dirs = []
    min_size_bytes = min_size_mb * (1024 * 1024)
    c_drive = Path("C:\\")
    
    # Common directories to skip (speeds up scan)
    skip_patterns = [
        r'$Recycle.Bin', 'System Volume Information', 'ProgramData', 
        'hiberfil.sys', 'pagefile.sys', 'swapfile.sys',
        'Windows\\System32', 'Windows\\WinSxS', '.git', '__pycache__'
    ]

    def should_skip(path_str):
        for pattern in skip_patterns:
            if pattern.lower() in path_str.lower():
                return True
        return False

    try:
        for root, dirs, filenames in os.walk(c_drive, topdown=True):
            # Filter out directories to skip
            dirs[:] = [d for d in dirs if not should_skip(os.path.join(root, d))]
            
            if should_skip(root):
                skipped_dirs.append(root)
                continue
            
            for filename in filenames:
                if len(files) >= max_results:
                    print(f"Reached max results limit ({max_results})")
                    break
                
                try:
                    filepath = Path(root) / filename
                    file_size = filepath.stat().st_size
                    
                    if file_size >= min_size_bytes:
                        size_mb = file_size / (1024 * 1024)
                        size_gb = file_size / (1024 * 1024 * 1024)
                        file_ext = filepath.suffix if filepath.suffix else 'No Extension'
                        
                        files.append({
                            'File Path': str(filepath),
                            'File Name': filename,
                            'Size MB': round(size_mb, 2),
                            'Size GB': round(size_gb, 3),
                            'File Type': file_ext,
                            'Modified': filepath.stat().st_mtime,
                        })
                        
                        if len(files) % 100 == 0:
                            print(f"  Found {len(files)} files so far...")
                except (PermissionError, OSError):
                    pass
            
            if len(files) >= max_results:
                break

    except Exception as e:
        print(f"Error during scan: {e}")

    # Sort by size descending
    files.sort(key=lambda x: x['Size MB'], reverse=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Large Files"

    total_size_gb = sum(f['Size GB'] for f in files)
    ws['A1'] = f"C: Drive Large Files Report - {len(files)} files found ({total_size_gb:.2f} GB)"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = f"Files larger than {min_size_mb}MB | Scanned: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws['A2'].font = Font(italic=True, size=10)

    headers = ['File Path', 'File Name', 'Size MB', 'Size GB', 'File Type', 'Modified']
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font

    # Color code by file size
    warning_fill = PatternFill("solid", fgColor="FFE699")  # Yellow for >1GB
    critical_fill = PatternFill("solid", fgColor="FF6B6B")  # Red for >5GB
    alt_fill = PatternFill("solid", fgColor="D9E1F2")  # Light blue

    for r_idx, file_data in enumerate(files, start=4):
        size_gb = file_data['Size GB']
        
        if size_gb > 5:
            fill = critical_fill
        elif size_gb > 1:
            fill = warning_fill
        else:
            fill = alt_fill if r_idx % 2 == 0 else PatternFill()
        
        for col, key in enumerate(headers, start=1):
            value = file_data[key]
            # Format datetime
            if key == 'Modified':
                from datetime import datetime as dt
                value = dt.fromtimestamp(value).strftime('%Y-%m-%d %H:%M:%S')
            
            cell = ws.cell(row=r_idx, column=col, value=value)
            cell.fill = fill
            if key in ['Size MB', 'Size GB']:
                cell.alignment = Alignment(horizontal="right")
            elif key in ['File Type']:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left", wrap_text=True)

    col_widths = [50, 30, 12, 10, 15, 20]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[chr(64+i)].width = w

    # Add summary sheet
    ws_summary = wb.create_sheet("Summary", 0)
    ws_summary['A1'] = "Large Files Summary"
    ws_summary['A1'].font = Font(bold=True, size=14)
    
    ws_summary['A3'] = "Metric"
    ws_summary['B3'] = "Value"
    for col in ['A', 'B']:
        ws_summary[f'{col}3'].font = Font(bold=True)
        ws_summary[f'{col}3'].fill = PatternFill("solid", fgColor="1F4E79")
        ws_summary[f'{col}3'].font = Font(bold=True, color="FFFFFF")

    summary_data = [
        ('Total Files Found', len(files)),
        ('Total Size (GB)', round(total_size_gb, 2)),
        ('Average File Size (MB)', round(sum(f['Size MB'] for f in files) / len(files), 2) if files else 0),
        ('Largest File (GB)', round(files[0]['Size GB'], 3) if files else 0),
        ('Largest File Name', files[0]['File Name'] if files else 'N/A'),
    ]
    
    row_num = 4
    for metric, value in summary_data:
        ws_summary[f'A{row_num}'] = metric
        ws_summary[f'B{row_num}'] = value
        row_num += 1
    
    # File type distribution
    ws_summary['A11'] = "File Type Distribution"
    ws_summary['A11'].font = Font(bold=True, size=12)
    
    ws_summary['A13'] = "File Type"
    ws_summary['B13'] = "Count"
    ws_summary['C13'] = "Total Size GB"
    for col in ['A', 'B', 'C']:
        ws_summary[f'{col}13'].font = Font(bold=True)
        ws_summary[f'{col}13'].fill = PatternFill("solid", fgColor="1F4E79")
        ws_summary[f'{col}13'].font = Font(bold=True, color="FFFFFF")
    
    file_types = {}
    for f in files:
        ftype = f['File Type']
        if ftype not in file_types:
            file_types[ftype] = {'count': 0, 'size': 0}
        file_types[ftype]['count'] += 1
        file_types[ftype]['size'] += f['Size GB']
    
    row_num = 14
    for ftype in sorted(file_types.keys(), key=lambda x: file_types[x]['size'], reverse=True):
        ws_summary[f'A{row_num}'] = ftype
        ws_summary[f'B{row_num}'] = file_types[ftype]['count']
        ws_summary[f'C{row_num}'] = round(file_types[ftype]['size'], 2)
        row_num += 1

    ws_summary.column_dimensions['A'].width = 25
    ws_summary.column_dimensions['B'].width = 15
    ws_summary.column_dimensions['C'].width = 15

    wb.save(output_path)
    print(f"\n[SUCCESS] Large files scan complete!")
    print(f"  Files found: {len(files)}")
    print(f"  Total size: {total_size_gb:.2f} GB")
    print(f"  Output: {output_path}")
