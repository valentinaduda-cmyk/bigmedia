"""
Timecode <-> frame conversions.

Library module — not wired to a CLI command yet, since a real "compute
runtime / trim / offset" command depends on which column(s) a given
workbook uses for in/out timecodes and at what frame rate. Import these
helpers directly for now; once we build the first timecode-based command
together, add a matching subcommand in cli.py.
"""
import math


def tc_to_frames(tc: str, fps: int = 25) -> int:
    h, m, s, f = map(int, tc.strip().split(':'))
    return ((h * 60 + m) * 60 + s) * fps + f


def frames_to_tc(frames: int, fps: int = 25) -> str:
    ff = frames % fps
    total_seconds = frames // fps
    ss = total_seconds % 60
    total_minutes = total_seconds // 60
    mm = total_minutes % 60
    hh = total_minutes // 60
    return f"{hh:02d}:{mm:02d}:{ss:02d}:{ff:02d}"


def frames_to_seconds_ceil(frames: int, fps: int = 25) -> int:
    return math.ceil(frames / fps)
