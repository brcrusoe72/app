from pathlib import Path
import subprocess


def test_no_merge_conflict_markers():
    repo = Path(__file__).resolve().parents[1]
    subprocess.check_call(["python", str(repo / "tools" / "check_merge_markers.py")], cwd=repo)
