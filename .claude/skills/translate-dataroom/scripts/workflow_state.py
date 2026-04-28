#!/usr/bin/env python3
"""
workflow_state.py — Helper for writing/updating .workflow_state.json

The orchestrator (skill / Claude) uses this to:
  1. INIT  the state file at job start with phase plan + estimates
  2. ADVANCE the current phase as it progresses
  3. COMPLETE when finished

The timer (timer.py --state ...) reads the file live and renders progress.

Usage:
  python workflow_state.py init <state_path> <english_folder> <job_name> --plan-json <plan>
  python workflow_state.py advance <state_path> <phase_name>
  python workflow_state.py progress <state_path> <fraction>
  python workflow_state.py complete <state_path>
  python workflow_state.py estimate <inventory_json>
      (returns a phase plan as JSON given a file inventory)
"""

import json
import sys
import os
import argparse
from datetime import datetime, timezone


# --- Estimation heuristics (seconds) ---
# These are calibrated from prior runs; adjust as more data accumulates.

def estimate_phases(inventory):
    """Build a phase plan from an inventory dict.

    inventory keys (all optional):
      - xls_count: int                 # files needing .xls→.xlsx conversion
      - copy_to_local_mb: float        # MB to copy to temp (if Drive Streaming)
      - translate_files: dict[str,int] # by extension, e.g. {"xlsx": 50, "pdf": 20}
      - translate_chars: int           # total JP characters (text-based)
      - translate_pdf_pages: int       # total PDF pages
      - mp4_minutes: float             # total audio minutes for transcription
      - deploy_to_drive_mb: float      # MB to copy back (if Drive Streaming)
      - xlsx_count: int                # for verification phase

    Returns list of phase dicts with name + estimated_seconds.
    """
    phases = []

    # Phase 1: convert .xls
    xls_n = inventory.get("xls_count", 0)
    if xls_n > 0:
        # Pure-Python fallback: ~1.5 sec/file. Excel COM tier: ~5 sec/file but may hang.
        # Budget for the slower path with a safety floor.
        phases.append({
            "name": "convert_xls",
            "file_count": xls_n,
            "estimated_seconds": max(15, xls_n * 3),
        })

    # Phase 2: copy to local temp (only if source is Drive Streaming)
    copy_mb = inventory.get("copy_to_local_mb", 0)
    if copy_mb > 0:
        # ~10 MB/sec is realistic for Drive Streaming reads
        phases.append({
            "name": "copy_to_local",
            "size_mb": copy_mb,
            "estimated_seconds": max(15, copy_mb / 10),
        })

    # Phase 3: translate (the main GCP work)
    tf = inventory.get("translate_files", {})
    chars = inventory.get("translate_chars", 0)
    pdf_pages = inventory.get("translate_pdf_pages", 0)

    # Per-file overhead (API round-trip): ~1 sec for text, ~5 sec for documents
    # Plus character-rate cost: GCP processes ~10K chars/sec for text translation
    text_files = sum(tf.get(k, 0) for k in ("xlsx", "xlsm", "csv", "txt", "md", "docx"))
    doc_files = sum(tf.get(k, 0) for k in ("pdf", "pptx"))

    text_seconds = text_files * 1.5 + chars / 10_000
    # PDFs: ~3 sec/page baseline + ~5 sec API round-trip per file
    pdf_seconds = doc_files * 5 + pdf_pages * 3

    translate_seconds = max(10, text_seconds + pdf_seconds)

    total_translate_files = text_files + doc_files
    phases.append({
        "name": "translate",
        "file_count": total_translate_files,
        "estimated_seconds": translate_seconds,
    })

    # Phase 4: transcribe MP4 (only if MP4 present)
    mp4_min = inventory.get("mp4_minutes", 0)
    if mp4_min > 0:
        # Whisper: ~10 sec processing per minute of audio + ~5 sec/min for translation
        phases.append({
            "name": "transcribe",
            "audio_minutes": mp4_min,
            "estimated_seconds": max(60, mp4_min * 15),
        })

    # Phase 5: deploy back to Drive (if Drive Streaming)
    deploy_mb = inventory.get("deploy_to_drive_mb", 0)
    if deploy_mb > 0:
        # Slower than read because of metadata + sync wait
        phases.append({
            "name": "deploy_to_drive",
            "size_mb": deploy_mb,
            "estimated_seconds": max(15, deploy_mb / 7),
        })

    # Phase 6: verify
    xlsx_count = inventory.get("xlsx_count", 0)
    if xlsx_count > 0:
        phases.append({
            "name": "verify",
            "file_count": xlsx_count,
            "estimated_seconds": max(5, xlsx_count * 0.2),
        })

    return phases


def cmd_estimate(inv_path):
    with open(inv_path, "r", encoding="utf-8") as f:
        inv = json.load(f)
    phases = estimate_phases(inv)
    total = sum(p["estimated_seconds"] for p in phases)
    print(json.dumps({"phases": phases, "total_seconds": total}, indent=2))


def cmd_init(state_path, english_folder, job_name, plan_json):
    plan = json.loads(plan_json) if plan_json else {"phases": [], "total_seconds": 0}
    state = {
        "job_name": job_name,
        "english_folder": english_folder,
        "started_at": datetime.now(timezone.utc).isoformat(),
        "phases": plan.get("phases", []),
        "current_phase_idx": 0,
        "current_phase_progress": 0.0,
        "complete": False,
    }
    os.makedirs(os.path.dirname(state_path), exist_ok=True)
    with open(state_path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)
    print(f"Wrote {state_path} with {len(state['phases'])} phases, total est {plan.get('total_seconds', 0):.0f}s")


def cmd_advance(state_path, phase_name):
    """Mark the named phase as the current one (advances index)."""
    with open(state_path, "r", encoding="utf-8") as f:
        state = json.load(f)
    found = False
    for i, p in enumerate(state["phases"]):
        if p["name"] == phase_name:
            state["current_phase_idx"] = i
            state["current_phase_progress"] = 0.0
            found = True
            break
    if not found:
        print(f"WARNING: phase {phase_name} not in plan; appending")
        state["phases"].append({"name": phase_name, "estimated_seconds": 30})
        state["current_phase_idx"] = len(state["phases"]) - 1
    with open(state_path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def cmd_progress(state_path, fraction):
    """Update progress within the current phase (0.0–1.0)."""
    with open(state_path, "r", encoding="utf-8") as f:
        state = json.load(f)
    state["current_phase_progress"] = max(0.0, min(1.0, float(fraction)))
    with open(state_path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def cmd_complete(state_path):
    with open(state_path, "r", encoding="utf-8") as f:
        state = json.load(f)
    state["complete"] = True
    state["current_phase_idx"] = len(state["phases"])
    state["completed_at"] = datetime.now(timezone.utc).isoformat()
    with open(state_path, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    cmd = sys.argv[1]

    if cmd == "estimate":
        cmd_estimate(sys.argv[2])
    elif cmd == "init":
        # init <state_path> <english_folder> <job_name> --plan-json <plan>
        state_path = sys.argv[2]
        english_folder = sys.argv[3]
        job_name = sys.argv[4]
        plan_json = None
        if "--plan-json" in sys.argv:
            i = sys.argv.index("--plan-json")
            plan_json = sys.argv[i + 1]
        cmd_init(state_path, english_folder, job_name, plan_json)
    elif cmd == "advance":
        cmd_advance(sys.argv[2], sys.argv[3])
    elif cmd == "progress":
        cmd_progress(sys.argv[2], sys.argv[3])
    elif cmd == "complete":
        cmd_complete(sys.argv[2])
    else:
        print(f"Unknown command: {cmd}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
