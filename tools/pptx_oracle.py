"""Run a python-pptx script and report the package it produced.

Invoked by spec/support/differential.rb. Reads a Python snippet on stdin,
executes it with `pptx` already imported and an `out` path bound, then prints a
JSON manifest of the resulting package: every zip entry with its canonicalised
content (XML is C14N-ised so attribute order and whitespace do not matter).
"""

from __future__ import annotations

import base64
import json
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


def canonicalize(name: str, data: bytes):
    if not name.lower().endswith((".xml", ".rels")):
        return {"kind": "binary", "sha": __import__("hashlib").sha256(data).hexdigest()}
    try:
        parser = etree.XMLParser(remove_blank_text=True)
        root = etree.fromstring(data, parser)
        # C14N 1.0 (not 2.0): it preserves unused namespace declarations, which are a
        # real fidelity difference, and it is what Nokogiri XML_C14N_1_0 emits.
        return {"kind": "xml", "c14n": etree.tostring(root, method="c14n").decode("utf-8")}
    except etree.XMLSyntaxError as exc:
        return {"kind": "malformed", "error": str(exc), "raw": base64.b64encode(data).decode()}


def manifest(path: Path):
    with zipfile.ZipFile(path) as zf:
        return {
            "entries": sorted(zf.namelist()),
            "parts": {n: canonicalize(n, zf.read(n)) for n in sorted(zf.namelist())},
        }


def main() -> int:
    script = sys.stdin.read()
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out.pptx"
        namespace = {"pptx": __import__("pptx"), "out": str(out), "Path": Path}
        try:
            exec(compile(script, "<oracle>", "exec"), namespace)  # noqa: S102
        except Exception as exc:  # noqa: BLE001
            json.dump({"ok": False, "error": f"{type(exc).__name__}: {exc}"}, sys.stdout)
            return 1
        if not out.exists():
            json.dump({"ok": False, "error": "script did not write to `out`"}, sys.stdout)
            return 1
        json.dump({"ok": True, **manifest(out)}, sys.stdout)
    return 0


if __name__ == "__main__":
    sys.exit(main())
