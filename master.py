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
    and writes them to an Excel file with one row per turn (prompt + response).

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
            ORDER BY s.created_at, t.turn_index
        """).fetchall()

        con.close()
        tmp.unlink(missing_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Copilot Sessions"

        headers = [
            "Session ID", "CWD", "Repository", "Branch",
            "Session Summary", "Session Created", "Session Updated",
            "Turn #", "Prompt (User)", "Response (Copilot)", "Turn Timestamp"
        ]

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
        for r_idx, row in enumerate(rows, start=2):
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
        print(f"[{now}] Exported {len(rows)} turns across {len(set(r['session_id'] for r in rows))} sessions -> {output_path}")

    if watch:
        print(f"Watch mode ON — refreshing every {interval_seconds}s. Press Ctrl+C to stop.")
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
