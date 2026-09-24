from __future__ import annotations

import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tests.xlsx_factory import make_xlsx


def main() -> None:
    samples = ROOT / "samples"
    samples.mkdir(parents=True, exist_ok=True)
    (samples / "config.json").write_text(
        json.dumps(
            {"required_columns": {"Data": ["ID", "Name"]}},
            indent=2,
        ) + "\n",
        encoding="utf-8",
    )

    make_xlsx(
        samples / "issue_rich.xlsx",
        [
            {
                "name": "Data",
                "rows": [
                    ["ID", "Name", "Amount", "Calc"],
                    [1, "Alpha", 10, {"formula": "MissingSheet!A1"}],
                    [1, "Alpha", 10, {"formula": "MissingSheet!A1"}],
                    [2, "", 20, {"error": "#REF!"}],
                ],
            },
            {
                "name": "HiddenData",
                "state": "hidden",
                "rows": [["Key"], ["x"]],
            },
        ],
    )

    make_xlsx(
        samples / "clean.xlsx",
        [
            {
                "name": "Data",
                "rows": [
                    ["ID", "Name", "Lookup"],
                    [1, "Alpha", {"formula": "Lookup!A1"}],
                    [2, "Beta", {"formula": "Lookup!A2"}],
                ],
            },
            {"name": "Lookup", "rows": [["Value"], [100], [200]]},
        ],
    )
    make_xlsx(
        samples / "quoted_refs.xlsx",
        [
            {
                "name": "Data",
                "rows": [
                    ["ID", "Calc"],
                    [1, {"formula": "'Valid Sheet'!A1"}],
                ],
            },
            {"name": "Valid Sheet", "rows": [["Value"], [10]]},
        ],
    )
    print("Generated 3 sample workbooks in", samples)


if __name__ == "__main__":
    main()
