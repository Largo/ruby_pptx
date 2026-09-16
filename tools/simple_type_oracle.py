"""Evaluate python-pptx simple-type conversions for a batch of cases.

Reads JSON [{"type": "ST_Angle", "op": "from_xml"|"to_xml", "value": ...}, ...]
on stdin and writes JSON results, so the Ruby specs can assert against the real
python-pptx behaviour rather than against a transcription of it.
"""

from __future__ import annotations

import json
import sys

from pptx.oxml import simpletypes


def describe(value):
    # Emu and friends subclass int; compare on the numeric value.
    if isinstance(value, bool):
        return {"ok": True, "kind": "bool", "value": value}
    if isinstance(value, int):
        return {"ok": True, "kind": "int", "value": value}
    if isinstance(value, float):
        return {"ok": True, "kind": "float", "value": value}
    return {"ok": True, "kind": "str", "value": str(value)}


def run(case):
    st = getattr(simpletypes, case["type"], None)
    if st is None:
        return {"ok": False, "error": "no such simple type"}
    try:
        if case["op"] == "from_xml":
            return describe(st.from_xml(case["value"]))
        return describe(st.to_xml(case["value"]))
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": type(exc).__name__}


def main() -> int:
    cases = json.load(sys.stdin)
    json.dump([run(c) for c in cases], sys.stdout)
    return 0


if __name__ == "__main__":
    sys.exit(main())
