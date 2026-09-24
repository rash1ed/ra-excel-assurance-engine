from __future__ import annotations

from pathlib import Path
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
LOG = ROOT / "docs" / "validation.txt"
RESULTS = ROOT / "samples" / "results"
RESULTS.mkdir(parents=True, exist_ok=True)


def run(label: str, args: list[str]) -> str:
    proc = subprocess.run(
        args,
        cwd=ROOT,
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
    )
    output = (proc.stdout or "") + (proc.stderr or "")
    output = output.replace(str(ROOT), "<PROJECT_ROOT>")
    output = output.replace(tempfile.gettempdir(), "<TEMP>")
    output = output.replace(str(Path.home()), "<USER_HOME>")
    block = f"=== {label} ===\n{output.rstrip()}\nEXIT_CODE={proc.returncode}\n\n"
    print(block, end="")
    if proc.returncode != 0:
        raise SystemExit(f"{label} failed with exit code {proc.returncode}")
    return block
def cli_args(name: str) -> list[str]:
    return [
        sys.executable,
        "-m",
        "ra_excel_assurance",
        f"samples/{name}.xlsx",
        "--config",
        "samples/config.json",
        "--report",
        f"samples/results/{name}_report.txt",
        "--json-report",
        f"samples/results/{name}_report.json",
        "--review-copy",
        f"samples/results/{name}_review_copy.xlsx",
    ]


def main() -> int:
    blocks: list[str] = []
    blocks.append(
        run(
            "UNIT TESTS",
            [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"],
        )
    )
    for sample in ["issue_rich", "clean", "quoted_refs"]:
        blocks.append(run(sample.upper(), cli_args(sample)))
    LOG.write_text("".join(blocks), encoding="utf-8")
    print(f"Validation log: {LOG}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
