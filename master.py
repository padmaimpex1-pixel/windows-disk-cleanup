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
        
        # URLs (check first)
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
