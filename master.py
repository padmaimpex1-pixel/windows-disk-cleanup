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
# MENU
# ---------------------------------------------------------------------------

TASKS = {
    "1": ("Export Copilot sessions to Excel (once)", lambda: export_copilot_sessions_to_excel()),
    "2": ("Export Copilot sessions to Excel (watch/continuous)", lambda: export_copilot_sessions_to_excel(watch=True, interval_seconds=60)),
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
