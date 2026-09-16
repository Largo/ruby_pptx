# Porting python-pptx to Ruby

Source of truth: **python-pptx 1.0.2** (MIT, Steve Canny). ~27.8k LOC of source,
~25k LOC of pytest units, 54 behave `.feature` files.

## Decisions

| Decision | Choice |
|---|---|
| Gem / repo / module | `ruby_pptx` / `Largo/ruby_pptx` / `Pptx` |
| Public API | **Ruby-idiomatic redesign** (not a 1:1 mirror) |
| Internal layers (`opc`, `oxml`) | Stay structurally close to python-pptx, so upstream fixes port across |
| Scope | Full port, module by module, including `chart/` |
| Tests | Differential against python-pptx + ported `.feature` scenarios; the mock-heavy pytest units are *not* ported |
| Deps | nokogiri (lxml), rubyzip (zipfile), hand-rolled image-header + TTF parsing (Pillow), minimal xlsx writer (XlsxWriter) |

### Two layers, one engine

The split that makes "idiomatic redesign" affordable:

- **Engine** (`Pptx::Opc`, `Pptx::Oxml`, parts, the shape/text object model) tracks
  python-pptx's structure closely. This is the hard, battle-tested part and we
  want to be able to read upstream commits as patches.
- **API** (everything a user touches) is redesigned for Ruby: `Enumerable`
  collections, `attr`-style accessors, keyword arguments instead of positional
  option soup, blocks for construction, no `get_or_add_x` leaking out.

### The one real architectural risk: `oxml/xmlchemy.py`

python-pptx declares its ~200 OOXML element classes with a metaclass DSL over
**lxml custom element classes** — lxml's parser is told to instantiate
`CT_Shape` (a real `ElementBase` subclass) whenever it meets `<p:sp>`. Node
identity and `getparent()` come free.

Nokogiri has no equivalent. The port uses a **wrapper layer**:

- `Pptx::Oxml::Element` wraps a `Nokogiri::XML::Node`.
- A tag registry maps Clark name → wrapper class (`"{...}sp"` → `Pptx::Oxml::Shapes::CT_Shape`).
- Each document carries an **identity map** (`Hash.compare_by_identity`, node → wrapper)
  so re-fetching the same child returns the *same* wrapper object and
  `equal?` comparisons behave as they do upstream.
- The declarative DSL becomes plain Ruby class macros, which read better than
  the Python original:

  ```ruby
  class CT_Shape < Element
    one_and_only_one :nvSpPr,  tag: "p:nvSpPr"
    zero_or_one      :spPr,    tag: "p:spPr", successors: %w[p:style p:txBody]
    zero_or_more     :ext,     tag: "p:extLst/p:ext"
    optional_attr    :macro,   type: XsdString
  end
  ```

This decision propagates through all ~4.1k LOC of `oxml/`, so it is settled
before anything else is written.

## Module inventory and milestones

LOC are python-pptx source lines, as a rough size signal.

### M1 — Foundation *(in progress)*

| Upstream | LOC | Ruby | Status |
|---|---:|---|---|
| `util.py` | 214 | `Pptx::Length` + opt-in `core_ext` | done |
| `oxml/ns.py` | 129 | `Pptx::Oxml::Ns` | done |
| `opc/packuri.py` | 109 | `Pptx::Opc::PackURI` | done |
| `exc.py` | 23 | `Pptx::Error` and friends | done |
| `oxml/xmlchemy.py` | 717 | `Pptx::Oxml::Element` + class-macro DSL | next |
| `oxml/simpletypes.py` | 740 | `Pptx::Oxml::SimpleTypes` | next |
| `enum/*` | 3087 | `Pptx::Enum::*` — mostly mechanical | next |

### M2 — OPC package (open/save round-trip)

`opc/package.py` (762), `opc/serialized.py` (296), `opc/oxml.py` (188),
`opc/constants.py` (331). Exit criterion: open a .pptx and save it byte-stable
against the differential harness.

### M3 — Presentation, parts, slides

`package.py` (222), `presentation.py` (113), `slide.py` (498),
`parts/*` (1090), `oxml/slide.py` (347), `oxml/presentation.py` (130),
`oxml/coreprops.py` (288).

### M4 — Shapes

`shapes/*` (3294, `shapetree.py` alone is 1190), `oxml/shapes/*`, `spec.py` (632),
`action.py` (270).

### M5 — Text and DrawingML

`text/text.py` (681), `oxml/text.py` (618), `dml/*` (880),
`text/fonts.py` (399, a TTF name-table parser — pure `struct`, ports cleanly),
`text/layout.py` (325).

### M6 — Tables, images, media

`table.py` (496), `oxml/table.py` (588), `parts/image.py` (275, needs our own
PNG/JPEG/GIF/BMP/TIFF header reader in place of Pillow), `media.py` (197).

### M7 — Charts

`chart/*` (5187; `xmlwriter.py` is 1840), `oxml/chart/*`. Needs a minimal
xlsx writer for the embedded chart workbook.

### M8 — Beyond python-pptx

See below.

## Beyond python-pptx: what to take from PptxGenJS

PptxGenJS is a **write-only generator**, not an object model — it cannot open an
existing deck. So it is not an alternative port target; it is a source of two
things.

**1. Authoring ergonomics.** Its one-call-per-object style is genuinely nicer
than python-pptx's `add_textbox(left, top, width, height)` positional
arguments. This belongs in the redesigned API layer, not in a second engine:

```ruby
Pptx.build do |deck|
  deck.slide(layout: :title_and_content) do |s|
    s.title = "Quarterly results"
    s.text "Revenue up 12%", at: [Pptx.inches(1), Pptx.inches(2)],
                             size: Pptx.pt(18), bold: true
    s.table rows, autopage: true
  end
end
```

**2. Five capability gaps, each verified absent from python-pptx 1.0.2:**

| Gap | Evidence | Milestone |
|---|---|---|
| Slide **sections** (`p14:sectionLst`) | no occurrence in the source | M8 |
| **Defining** a slide master/layout in code | only `add_slide(existing_layout)` exists | M8 |
| **Combo charts** and secondary axes | multi-plot charts can be *read*; `add_chart` only ever writes one plot | M7 |
| **SVG** image parts | no occurrence in the source | M6 |
| Table **autopaging** across slides | not a python-pptx concept | M6 |

Deliberately *not* taken: HTML-table import and YouTube embeds (separate
add-on gems if wanted), and Blob/base64/stream export, which in Ruby is just
writing to an `IO`.

## Testing

`spec/support/differential.rb` drives the same operation through python-pptx
(importable on this box) and through the gem, then diffs the resulting package:
entry list, then canonicalised XML per part. Failures print a part-level diff.
The 54 `.feature` files are ported as acceptance specs; their Gherkin
translates directly even though the step definitions are rewritten against the
Ruby API.
