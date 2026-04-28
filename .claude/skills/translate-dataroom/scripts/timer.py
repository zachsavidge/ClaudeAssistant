#!/usr/bin/env python3
"""
timer.py - Live progress-aware timer for the entire data room translation workflow.

Tracks ALL phases (not just GCP translation):
  1. convert_xls    (xls → xlsx via Excel COM or pure Python)
  2. copy_to_local  (Drive Streaming → local temp, if applicable)
  3. translate      (the GCP work — main phase)
  4. transcribe     (MP4/audio via Whisper, if applicable)
  5. deploy_to_drive (local temp → Drive English folder, if applicable)
  6. verify         (Excel ZIP integrity + formula check)

Reads a state file written by the orchestrator (claude/skill) describing the
phases, their estimated durations, and live progress per phase. Adjusts ETA
based on actual rate vs estimate.

Usage:
    # Phase-aware mode (preferred):
    python timer.py --state <english_folder>/.workflow_state.json

    # Legacy mode (translation phase only):
    python timer.py <total_files> <english_folder>
"""

import json
import time
import sys
import os
import argparse
from datetime import datetime, timezone


def parse_args():
    p = argparse.ArgumentParser(add_help=False)
    p.add_argument("--state", help="Path to .workflow_state.json")
    p.add_argument("positional", nargs="*")
    args = p.parse_args()
    return args


def load_state(state_path):
    """Load workflow state. Returns None if missing/unreadable."""
    try:
        with open(state_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def get_translation_completed(english_folder):
    """Read .translation_state.json for live translation progress."""
    f = os.path.join(english_folder, ".translation_state.json")
    try:
        with open(f, "r", encoding="utf-8") as fp:
            state = json.load(fp)
        return len(state.get("completed", []))
    except Exception:
        return 0


def format_time(seconds):
    seconds = max(0, int(seconds))
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    if h > 0:
        return f"{h:02d}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"


def render_bar(pct, length=40):
    pct = max(0, min(100, pct))
    filled = int(length * pct / 100)
    return "#" * filled + "-" * (length - filled)


PHASE_LABELS = {
    "convert_xls":      "1. Convert .xls files",
    "copy_to_local":    "2. Copy to local temp",
    "translate":        "3. Translate (GCP)",
    "transcribe":       "4. Transcribe MP4 (Whisper)",
    "deploy_to_drive":  "5. Deploy to Drive",
    "verify":           "6. Verify integrity",
}


def auto_progress(state, english_folder):
    """For phases with auto-detection, compute live progress.

    Currently auto-detects:
      - 'translate' phase: counts .translation_state.json
    """
    phases = state.get("phases", [])
    cur_idx = state.get("current_phase_idx", 0)
    if cur_idx >= len(phases):
        return state

    cur = phases[cur_idx]
    if cur.get("name") == "translate":
        completed = get_translation_completed(english_folder)
        total = cur.get("file_count") or state.get("total_files", 1)
        if total > 0:
            cur["live_progress"] = min(1.0, completed / total)
            cur["live_completed"] = completed
            cur["live_total"] = total
    return state


def render(state, english_folder, started_at_epoch):
    """Render the full workflow status to stdout."""
    os.system("cls" if os.name == "nt" else "clear")

    elapsed = time.time() - started_at_epoch
    phases = state.get("phases", [])
    cur_idx = state.get("current_phase_idx", 0)

    # Calculate total estimated time and remaining time
    total_est = sum(p.get("estimated_seconds", 0) for p in phases)
    elapsed_in_phases = 0
    for i, p in enumerate(phases):
        if i < cur_idx:
            elapsed_in_phases += p.get("estimated_seconds", 0)
        elif i == cur_idx:
            prog = p.get("live_progress", state.get("current_phase_progress", 0))
            elapsed_in_phases += p.get("estimated_seconds", 0) * prog

    # Adjust estimates based on actual elapsed vs expected
    if elapsed_in_phases > 0 and elapsed > 0:
        adjust = elapsed / elapsed_in_phases
        adjust = max(0.5, min(3.0, adjust))  # clamp drift
    else:
        adjust = 1.0

    remaining_est = (total_est - elapsed_in_phases) * adjust
    overall_pct = (elapsed_in_phases / total_est * 100) if total_est > 0 else 0

    # Header
    print("=" * 65)
    print("   DATAROOM TRANSLATION — END-TO-END PROGRESS")
    print("=" * 65)
    print()
    job = state.get("job_name", "")
    if job:
        print(f"   Job: {job}")
    print(f"   Elapsed:   {format_time(elapsed)}")
    print(f"   ETA:       {format_time(remaining_est)}")
    print(f"   Total est: {format_time(total_est * adjust)}")
    print()
    print(f"   Overall: [{render_bar(overall_pct)}] {overall_pct:5.1f}%")
    print()
    print("-" * 65)

    # Per-phase status
    for i, p in enumerate(phases):
        name = p.get("name", "?")
        label = PHASE_LABELS.get(name, name)
        est = p.get("estimated_seconds", 0)

        if i < cur_idx:
            status = "DONE"
            symbol = "[X]"
            extra = ""
        elif i == cur_idx:
            symbol = "[>]"
            prog = p.get("live_progress", state.get("current_phase_progress", 0))
            status = f"{prog * 100:5.1f}%"
            if "live_completed" in p:
                extra = f"   ({p['live_completed']}/{p['live_total']} files)"
            elif "file_count" in p:
                extra = f"   ({p['file_count']} files)"
            elif "size_mb" in p:
                extra = f"   ({p['size_mb']:.0f} MB)"
            else:
                extra = ""
        else:
            symbol = "[ ]"
            status = "queued"
            if "file_count" in p:
                extra = f"   (~{p['file_count']} files)"
            elif "size_mb" in p:
                extra = f"   (~{p['size_mb']:.0f} MB)"
            else:
                extra = ""

        print(f"   {symbol} {label:<32} {status:>7}  est {format_time(est)}{extra}")

    print("-" * 65)
    print()

    # Status line
    cur_phase = phases[cur_idx] if cur_idx < len(phases) else None
    if state.get("complete"):
        print("   STATUS: ✓ COMPLETE")
    elif cur_phase and cur_phase.get("name") == "translate":
        completed = cur_phase.get("live_completed", 0)
        total = cur_phase.get("live_total", 0)
        if completed == 0 and elapsed > 10:
            print("   STATUS: Waiting for translation script to start...")
        else:
            print(f"   STATUS: Translating ({completed}/{total} files)")
    else:
        print(f"   STATUS: {PHASE_LABELS.get(cur_phase['name'], cur_phase['name']) if cur_phase else 'preparing'}")

    print()
    print("=" * 65)


def beep_done():
    """Cross-platform completion beep. 5 beeps on Windows, terminal bell elsewhere."""
    if sys.platform == "win32":
        try:
            import winsound
            for _ in range(5):
                winsound.Beep(1000, 500)
                time.sleep(0.3)
            return
        except Exception:
            pass
    elif sys.platform == "darwin":
        # macOS: use afplay with the system sound, fallback to terminal bell
        import subprocess
        for _ in range(5):
            try:
                subprocess.run(
                    ["afplay", "/System/Library/Sounds/Glass.aiff"],
                    timeout=2, check=False,
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                )
            except Exception:
                print("\a", end="", flush=True)
            time.sleep(0.3)
        return
    # Linux / fallback: terminal bell
    for _ in range(5):
        print("\a", end="", flush=True)
        time.sleep(0.3)


def main_phase_aware(state_path):
    """Run in phase-aware mode."""
    # Wait for state file to exist (orchestrator may take a moment to write it)
    deadline = time.time() + 30
    while not os.path.exists(state_path) and time.time() < deadline:
        time.sleep(1)
    if not os.path.exists(state_path):
        print(f"State file never appeared: {state_path}")
        sys.exit(1)

    started = time.time()
    english_folder = None

    try:
        while True:
            state = load_state(state_path)
            if state is None:
                time.sleep(2)
                continue

            # Pick up english_folder once known
            if english_folder is None:
                english_folder = state.get("english_folder", "")

            # Auto-update live progress where possible
            state = auto_progress(state, english_folder)

            # Check completion marker
            complete_marker = os.path.join(english_folder, ".translation_complete") if english_folder else None
            if state.get("complete") or (complete_marker and os.path.exists(complete_marker)):
                # Mark all phases done
                state["complete"] = True
                state["current_phase_idx"] = len(state.get("phases", []))
                render(state, english_folder, started)
                print("   Press Ctrl+C to close.")
                beep_done()
                time.sleep(999999)
                break

            render(state, english_folder, started)
            time.sleep(2)

    except KeyboardInterrupt:
        print("\nTimer closed.")


def main_legacy(total_files, english_folder):
    """Legacy mode: just track translation phase from .translation_state.json."""
    fake_state = {
        "job_name": os.path.basename(english_folder),
        "english_folder": english_folder,
        "phases": [
            {"name": "translate", "estimated_seconds": total_files * 5,
             "file_count": total_files}
        ],
        "current_phase_idx": 0,
        "current_phase_progress": 0,
    }
    started = time.time()
    marker = os.path.join(english_folder, ".translation_complete")

    try:
        while True:
            fake_state = auto_progress(fake_state, english_folder)
            if os.path.exists(marker):
                fake_state["complete"] = True
                fake_state["current_phase_idx"] = 1
                render(fake_state, english_folder, started)
                print("   Press Ctrl+C to close.")
                beep_done()
                time.sleep(999999)
                break
            render(fake_state, english_folder, started)
            time.sleep(2)
    except KeyboardInterrupt:
        print("\nTimer closed.")


def main():
    args = parse_args()
    if args.state:
        main_phase_aware(args.state)
    elif len(args.positional) >= 2:
        main_legacy(int(args.positional[0]), args.positional[1])
    else:
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
