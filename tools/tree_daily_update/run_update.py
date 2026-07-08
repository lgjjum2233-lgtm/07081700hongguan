from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from audit_workbooks import audit, json_default
from postprocess_daily import default_output_path, load_config, postprocess_daily
from sync_tree import default_output_path as default_tree_output_path
from sync_tree import sync_tree


def newest_file(workspace: Path, pattern: str) -> Path | None:
    files = [p for p in workspace.glob(pattern) if p.is_file() and not p.name.startswith("~$")]
    return max(files, key=lambda p: p.stat().st_mtime) if files else None


def run_refresh(refresh_script: Path, workbook: Path, wait_seconds: int, output: Path | None = None) -> dict[str, Any]:
    cmd = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(refresh_script),
        "-WorkbookPath",
        str(workbook),
        "-WaitSeconds",
        str(wait_seconds),
    ]
    if output:
        cmd.extend(["-OutputPath", str(output)])
    proc = subprocess.run(cmd, text=True, encoding="utf-8", errors="replace", capture_output=True, check=False)
    return {
        "command": " ".join(cmd),
        "returncode": proc.returncode,
        "stdout": (proc.stdout or "").strip(),
        "stderr": (proc.stderr or "").strip(),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Low-token TREE daily update guardrail runner.")
    parser.add_argument("--daily", help="日报 workbook path. Defaults to newest *日报*.xlsx in workspace.")
    parser.add_argument("--tree", help="TREE workbook path. Defaults to newest *TREE*.xlsx in workspace.")
    parser.add_argument("--config", default=str(Path(__file__).with_name("config.json")))
    parser.add_argument("--postprocess-daily", action="store_true", help="Rewrite front chart caches before audit.")
    parser.add_argument("--daily-output", help="Output path for postprocessed daily workbook.")
    parser.add_argument("--refresh-daily", action="store_true", help="Refresh daily workbook through WPS/Wind before postprocess/audit.")
    parser.add_argument("--refresh-output", help="Optional SaveCopyAs output for WPS refresh.")
    parser.add_argument("--wait-seconds", type=int, default=120)
    parser.add_argument("--skip-deep", action="store_true", help="Skip the heavier TREE deep audit.")
    parser.add_argument("--summary-output", help="Optional JSON summary output path.")
    parser.add_argument("--sync-tree", action="store_true", help="Synchronize TREE from the final daily workbook before audit.")
    parser.add_argument("--tree-output", help="Output path for synchronized TREE workbook.")
    parser.add_argument("--previous-tree", help="Previous TREE baseline for same-day marginal-change coloring.")
    args = parser.parse_args()

    config = load_config(Path(args.config))
    workspace = Path(config["workspace"])
    daily_path = Path(args.daily) if args.daily else newest_file(workspace, "*日报*.xlsx")
    tree_path = Path(args.tree) if args.tree else newest_file(workspace, "*TREE*.xlsx")
    if daily_path is None:
        raise FileNotFoundError("No daily workbook found. Pass --daily.")
    if tree_path is None:
        raise FileNotFoundError("No TREE workbook found. Pass --tree.")

    result: dict[str, Any] = {
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "input_daily": str(daily_path),
        "input_tree": str(tree_path),
    }
    previous_tree_for_colors = Path(args.previous_tree) if args.previous_tree else tree_path

    if args.refresh_daily:
        refresh_script = Path(__file__).with_name("refresh_daily.ps1")
        refresh_output = Path(args.refresh_output) if args.refresh_output else None
        refresh_result = run_refresh(refresh_script, daily_path, args.wait_seconds, refresh_output)
        result["refresh_daily"] = refresh_result
        if refresh_result["returncode"] != 0:
            result["summary"] = {"pass": False, "blocked_at": "refresh_daily"}
            print(json.dumps(result, ensure_ascii=False, indent=2, default=json_default))
            sys.exit(1)
        if refresh_output:
            daily_path = refresh_output

    if args.postprocess_daily:
        output_path = Path(args.daily_output) if args.daily_output else default_output_path(daily_path, config)
        post_result = postprocess_daily(daily_path, output_path, config)
        result["postprocess_daily"] = post_result
        daily_path = output_path
        if post_result["chart_last_current_mismatch_count"] != 0 or post_result["charts_not_updated"] != 0:
            result["summary"] = {"pass": False, "blocked_at": "postprocess_daily"}
            print(json.dumps(result, ensure_ascii=False, indent=2, default=json_default))
            sys.exit(1)

    if args.sync_tree:
        tree_output = Path(args.tree_output) if args.tree_output else default_tree_output_path(tree_path)
        sync_result = sync_tree(daily_path, tree_path, tree_output, previous_tree_for_colors, config)
        result["sync_tree"] = sync_result
        tree_path = tree_output
        color_summary = sync_result.get("colors", {})
        if color_summary.get("bad_color_count") or color_summary.get("extra_color_count"):
            result["summary"] = {"pass": False, "blocked_at": "sync_tree_colors"}
            print(json.dumps(result, ensure_ascii=False, indent=2, default=json_default))
            sys.exit(1)

    audit_result = audit(
        daily_path,
        tree_path,
        config,
        skip_deep=args.skip_deep,
        previous_tree=previous_tree_for_colors if tree_path else None,
    )
    result["audit"] = audit_result
    result["final_daily"] = str(daily_path)
    result["final_tree"] = str(tree_path)
    result["summary"] = {"pass": bool(audit_result.get("summary", {}).get("pass"))}

    text = json.dumps(result, ensure_ascii=False, indent=2, default=json_default)
    if args.summary_output:
        Path(args.summary_output).write_text(text, encoding="utf-8")
    print(text)
    sys.exit(0 if result["summary"]["pass"] else 2)


if __name__ == "__main__":
    main()
