#!/usr/bin/env python3
"""Open a .pptx in Microsoft PowerPoint (macOS), accept Repair, save to a new path."""

from __future__ import annotations

import subprocess
import time
from pathlib import Path


def powerpoint_repair_and_save(src: Path, dst: Path, *, timeout_s: int = 90) -> None:
    src = src.resolve()
    dst = dst.resolve()
    if not src.is_file():
        raise FileNotFoundError(src)

    script = f'''
set srcFile to POSIX file "{src}"
set dstFile to POSIX file "{dst}"

tell application "Microsoft PowerPoint"
    activate
    open srcFile
end tell
set clickedRepair to false
repeat with attempt from 1 to 12
    delay 2
    tell application "System Events"
        tell process "Microsoft PowerPoint"
            repeat with w in windows
                repeat with b in buttons of w
                    try
                        set bn to name of b as text
                        if bn contains "복구" or bn contains "Repair" then
                            click b
                            set clickedRepair to true
                            exit repeat
                        end if
                    end try
                end repeat
                if clickedRepair then exit repeat
                try
                    repeat with s in sheets of w
                        repeat with b in buttons of s
                            try
                                set bn to name of b as text
                                if bn contains "복구" or bn contains "Repair" then
                                    click b
                                    set clickedRepair to true
                                    exit repeat
                                end if
                            end try
                        end repeat
                    end repeat
                end try
                if clickedRepair then exit repeat
            end repeat
        end tell
    end tell
    if clickedRepair then exit repeat
end repeat

delay 3

tell application "Microsoft PowerPoint"
    save active presentation in dstFile
    close active presentation saving no
end tell
'''
    proc = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
        timeout=timeout_s,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"PowerPoint repair failed (code {proc.returncode}): {proc.stderr or proc.stdout}"
        )
    if not dst.is_file():
        raise RuntimeError(f"PowerPoint did not write output: {dst}")
    time.sleep(0.5)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("src", type=Path)
    parser.add_argument("dst", type=Path)
    args = parser.parse_args()
    powerpoint_repair_and_save(args.src, args.dst)
    print(f"Repaired copy saved: {args.dst}")


if __name__ == "__main__":
    main()
