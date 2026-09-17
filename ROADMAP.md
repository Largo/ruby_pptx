# Roadmap

State as of the M8 commit: 402 specs, green on Ruby 3.3, 3.4 and 4.0.
See [PORTING.md](PORTING.md) for what each milestone covered and why.

Two tracks: making the API read like Ruby, and finishing the port. The first
is smaller and should go first — it changes signatures, and every week of
delay makes that more disruptive.

---

## A. API idiom — **done**

Everything in this section shipped. Kept here as the record of what changed
and why; see the sub-sections for the reasoning.



The project's stated decision was a **Ruby-idiomatic redesign** over a faithful
engine. The engine held up; the public API drifted back toward python-pptx's
shape in places. Findings below are from reading the actual public surface, not
from recollection.

### A1. Geometry is positional in the core API — the biggest gap

```ruby
# today
shapes.add_shape(type, left, top, width, height)
shapes.add_table(rows, cols, left, top, width, height)
shapes.add_chart(type, left, top, width, height, data)   # data last, after geometry
shapes.add_picture(file, left, top, width:, height:)     # mixed: positional position, keyword size
```

Four positional lengths in a row is exactly what the redesign was meant to
avoid, and `add_picture` is already inconsistent with its neighbours. The
builder (`Pptx.build`) has the right shape; the core should match it.

```ruby
# proposed
shapes.add_shape(type, at: [left, top], size: [width, height])
shapes.add_table(rows, cols, at:, size:)
shapes.add_chart(type, data, at:, size:)
shapes.add_picture(file, at:, width: nil, height: nil)
```

Breaking. Worth doing before anyone depends on it. Keep the positional forms
as deprecated aliases for one release if that seems kind.

**Size**: half a day including specs. The builder already proves the two forms
agree, so the self-differential spec protects the change.

### A2. `Presentation.open` silently ignores a block

```ruby
Pptx::Presentation.open(path) { |prs| ... }   # block never runs, no error
```

Every Rubyist reaches for this by reflex because of `File.open`. Silently
doing nothing is the worst of the three options. Either yield and return the
block's value, or raise. **Size**: an hour.

### A3. Enumerations are classes named in SCREAMING_SNAKE

```ruby
Pptx::Enum::MSO_SHAPE::ROUNDED_RECTANGLE
shapes.add_shape(:OVAL, ...)    # works
shapes.add_shape(:oval, ...)    # ArgumentError
```

`Pptx::Enum::MSO_AUTO_SHAPE_TYPE` is a class whose name is not CamelCase, which
no Ruby style guide permits. It was a deliberate fidelity choice — the names
match the MS API and python-pptx — and there is real value in that for anyone
porting code or reading Microsoft's docs.

The cheap fix keeps both: accept `:oval` and `:rounded_rectangle` anywhere an
enum member is taken, by downcasing on lookup. That is non-breaking and makes
the common path read like Ruby. Renaming the classes themselves
(`Pptx::Enum::AutoShape`) is a larger, more debatable change; recommend
deferring it and revisiting only if the symbol path proves insufficient.

**Size**: symbol tolerance, two hours. Class renaming, a day, and not
recommended yet.

### A4. Boolean readers missing `?`

```ruby
table.first_row      #=> true
table.horz_banding   #=> true
```

Genuine booleans should be predicates: `first_row?`, `last_col?`,
`banded_rows?`. Also worth spelling out the abbreviations — `horz_banding`
is python-pptx's name, not a Ruby one.

Note this does *not* apply to `Font#bold`, which is tri-state: `nil` means
"inherited". A `bold?` returning `nil` would be worse than `bold`. Leave it.

**Size**: an hour, plus keeping the writers as `first_row=`.

### A5. `has_` prefixes are inconsistent

`Slide#has_notes_slide?` and `Chart#has_legend?` sit beside `text_frame?`,
`chart?`, `table?`, `placeholder?`. Pick one — the un-prefixed form, to match
the majority — and alias the old names. **Size**: half an hour.

### A6. Internal plumbing is public

- `Table#notify_height_changed` / `#notify_width_changed` are callbacks a row
  or column makes when resized. They are implementation, not API.
- `ChartData#categories_ref`, `#series_name_ref`, `#series_values_ref`,
  `#column_letter`, `#xlsx_blob` are worksheet plumbing the XML writer needs.

Neither should appear in the documented surface. Move behind a
`Pptx::Internal` module or mark them `@api private` and exclude from docs.
**Size**: an hour.

### A7. Small conveniences a Rubyist will expect

| Missing | Why |
|---|---|
| `table[row, col]` | alias for `cell(row, col)` |
| `2 * length` | `Length#coerce`; `length * 2` works, the reverse raises |
| `Presentation#to_blob` | `save` takes an IO, but a blob is the common want |
| `Slides#each_with_index` etc. | already works via Enumerable — no action |

**Size**: two hours together.

### What is already right

Worth recording so it does not get "fixed": Enumerable collections; `[]`
returning `nil` while `fetch` raises; frozen `Comparable` value objects;
`Length` arithmetic; `clear` without a bang (matching `Array#clear`);
`Sections#delete` returning the deleted section; `element` as a documented
escape hatch; and the oxml layer keeping schema-exact camelCase, which is
internal and deliberate.

**Verdict**: the API is close, and the shape is right. A1 and A2 are the two
that would actually make a reviewer wince. Everything else is polish.

**Outcome**: A1-A7 are all done, except that `Length#coerce` (part of A7) was
deliberately *not* added. Ruby's coerce protocol cannot see which operator is
being applied, so the pair it returns must be right for all of them. Returning
`[self, other]` would make `2 * length` work but silently turn `2 - length`
into `length - 2`, the wrong sign. A `TypeError` the caller fixes by writing
`length * 2` beats an answer that is quietly negative. The reasoning is
recorded in `lib/pptx/length.rb` so nobody "fixes" it later.

---

## B. Finishing the port

Remaining upstream code, measured:

| Area | Upstream LOC | Notes |
|---|---:|---|
| Chart formatting — axes, plots, data labels, legend, marker | ~1,530 object model + ~1,250 oxml | Charts can be created but not restyled. Largest remaining item. |
| Chart families — area, radar, XY, bubble | ~900 of `xmlwriter.py` | Same pattern as the four already done; mechanical. |
| Freeform shape building | 337 | Self-contained. |
| Connectors and groups | 366 | Read works; creating and manipulating does not. |
| Hyperlinks / click actions (`action.py`) | 323 | Wanted more often than charts, in practice. |
| Video (`media.py`, `parts/media.py`) | 234 | |
| `fit_text` — TTF parsing and line layout | 724 | Only `TextFrame#fit_text` needs it. |
| Table cell merging | part of `table.py` | Small. |

**Beyond python-pptx, still open** (from M8):

- **SVG images** — blocked on a design decision, not effort. PowerPoint wants a
  raster fallback in `r:embed` beside the SVG; this gem has no rasterizer. The
  honest signature is `add_picture(svg, fallback: png)`. Decide, then it is a
  day.
- **Defining a slide master in code** — needs a theme part with colour, font and
  format schemes plus `p:txStyles` before a layout can inherit. Genuinely large;
  starting from a template remains the practical route.
- **Combo charts / secondary axes** — depends on chart formatting above.

**Suggested order**: hyperlinks, then cell merging, then the remaining chart
families, then chart formatting. That front-loads what people actually reach
for and leaves the biggest item last.

---

## C. Before a release

- [ ] Decide the public API (section A) — signatures should settle first.
- [ ] `CHANGELOG.md`.
- [ ] Decide the version. `0.1.0` is honest for what works; `1.0` would claim
      completeness the roadmap above contradicts.
- [ ] Document the supported subset prominently, so nobody discovers by
      crashing that charts cannot be restyled.
- [ ] RuboCop, with a config that permits the oxml layer's camelCase.
- [ ] YARD docs; the codebase is already commented for it.
- [ ] Decide whether to publish. `ruby_pptx` is free on RubyGems.
