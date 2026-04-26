"""One-off: init repo and push to GitHub. Run: python _git_push_helper.py"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
REMOTE = "git@github.com:dangkhoipham80/convert_test_report_UT.git"

def run(cmd: list[str]) -> None:
    print("+", " ".join(cmd), flush=True)
    r = subprocess.run(cmd, cwd=ROOT, text=True)
    if r.returncode != 0:
        sys.exit(r.returncode)

def main() -> None:
    os.chdir(ROOT)
    if not os.path.isdir(os.path.join(ROOT, ".git")):
        run(["git", "init"])
    run(["git", "add", "."])
    # commit if there are changes
    st = subprocess.run(
        ["git", "status", "--porcelain"],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    if st.stdout.strip():
        run(["git", "commit", "-m", "Initial commit: test report / matrix conversion tools"])
    run(["git", "branch", "-M", "main"])
    # remote: add or set-url
    rr = subprocess.run(
        ["git", "remote", "get-url", "origin"],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    if rr.returncode != 0:
        run(["git", "remote", "add", "origin", REMOTE])
    else:
        run(["git", "remote", "set-url", "origin", REMOTE])
    run(["git", "push", "-u", "origin", "main"])
    print("Done.", flush=True)

if __name__ == "__main__":
    main()
