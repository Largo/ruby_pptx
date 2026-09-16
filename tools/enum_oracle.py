"""Dump every python-pptx enumeration as JSON, for the Ruby specs to check against."""

from __future__ import annotations

import importlib
import inspect
import json
import sys

from pptx.enum.base import BaseEnum, BaseXmlEnum

MODULES = ["action", "chart", "dml", "lang", "shapes", "text"]


def main() -> int:
    out: dict[str, dict] = {}
    for mod_name in MODULES:
        module = importlib.import_module("pptx.enum." + mod_name)
        for name, obj in vars(module).items():
            if not (inspect.isclass(obj) and issubclass(obj, (BaseEnum, BaseXmlEnum))):
                continue
            if obj in (BaseEnum, BaseXmlEnum) or obj.__module__ != module.__name__:
                continue
            out[name] = {
                "canonical": obj.__name__,
                "xml_mapped": issubclass(obj, BaseXmlEnum),
                "members": [
                    {
                        "name": m.name,
                        "value": int(m.value),
                        "xml": getattr(m, "xml_value", None) or None,
                    }
                    for m in obj
                ],
                # Several MS API members can share one XML value (the API
                # distinguishes callout 3 from callout 4; OOXML does not), so
                # record what from_xml actually resolves each value to.
                "from_xml": {
                    m.xml_value: obj.from_xml(m.xml_value).name
                    for m in obj
                    if getattr(m, "xml_value", None)
                },
            }
    json.dump(out, sys.stdout)
    return 0


if __name__ == "__main__":
    sys.exit(main())
