"""Serialize a python-pptx core-properties part after applying operations.

Reads JSON {"ops": [[name, value], ...]} on stdin, applies them to a freshly
created CT_CoreProperties, and writes the resulting part XML, so the Ruby specs
can compare against real python-pptx output for a part that does not yet exist
in any package.
"""

from __future__ import annotations

import datetime as dt
import json
import sys

from pptx.opc.oxml import serialize_part_xml
from pptx.oxml.coreprops import CT_CoreProperties

SETTERS = {
    "title": "title_text",
    "author": "author_text",
    "keywords": "keywords_text",
    "last_modified_by": "lastModifiedBy_text",
    "revision": "revision_number",
    "created": "created_datetime",
    "modified": "modified_datetime",
    "last_printed": "lastPrinted_datetime",
}


def main() -> int:
    payload = json.load(sys.stdin)
    element = CT_CoreProperties.new_coreProperties()
    for name, value in payload["ops"]:
        if name in ("created", "modified", "last_printed"):
            value = dt.datetime(*value)
        setattr(element, SETTERS[name], value)
    sys.stdout.write(serialize_part_xml(element).decode("utf-8"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
