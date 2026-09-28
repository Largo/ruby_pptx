"""Dump python-pptx's public API surface as JSON: class -> public member names.

Consumed by tools/api_audit.rb, which checks every member against this gem.
Internal layers (oxml, opc, parts, enum) are skipped; they are ported
structurally rather than as API.
"""

from __future__ import annotations

import importlib
import inspect
import json
import pkgutil
import sys

import pptx

SKIP = (".oxml", ".opc", ".parts", ".enum", "pptx.spec", "pptx.exc", "pptx.util")


def main() -> int:
    out = {}
    for mod in pkgutil.walk_packages(pptx.__path__, "pptx."):
        if any(part in mod.name for part in SKIP):
            continue
        module = importlib.import_module(mod.name)
        for name, cls in inspect.getmembers(module, inspect.isclass):
            if cls.__module__ != mod.name:
                continue
            members = sorted(
                n for n, _ in inspect.getmembers(cls)
                if not n.startswith("_") and n not in dir(object) and n not in ("count", "index")
            )
            if members:
                out[f"{mod.name.removeprefix('pptx.')}.{name}"] = members
    json.dump({"version": pptx.__version__, "classes": out}, sys.stdout, indent=1, sort_keys=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
