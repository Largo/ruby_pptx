"""Report Pillow's view of image files, for the Ruby image-header specs.

Reads a JSON list of paths on stdin, writes format, pixel size and the dpi
python-pptx would derive, so the hand-written header parser can be checked
against the library it replaces.
"""

from __future__ import annotations

import json
import sys

from PIL import Image

from pptx.parts.image import Image as PptxImage


def main() -> int:
    paths = json.load(sys.stdin)
    out = {}
    for path in paths:
        with Image.open(path) as im:
            fmt, size = im.format, list(im.size)
        pptx_image = PptxImage.from_file(path)
        out[path] = {
            "format": fmt,
            "size": size,
            "dpi": list(pptx_image.dpi),
            "ext": pptx_image.ext,
            "content_type": pptx_image.content_type,
            "sha1": pptx_image.sha1,
        }
    json.dump(out, sys.stdout)
    return 0


if __name__ == "__main__":
    sys.exit(main())
