# Changelog

All notable changes to this gem are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html). Before 1.0 a minor
version may change the public API.

## [Unreleased]

### Added

- A pure-Ruby XML backend built on REXML, used automatically where Nokogiri
  cannot be loaded -- notably ruby.wasm. It writes the same packages as the
  Nokogiri backend; CI runs every spec under both, and builds, saves and
  rereads a deck inside ruby.wasm. `Pptx.xml_backend` reports which is in
  use, and `RUBY_PPTX_XML_BACKEND` chooses one explicitly.
- `rexml` is now a runtime dependency. Nokogiri remains one too.

### Changed

- `require "ruby_pptx"` no longer requires Nokogiri directly; the XML
  backend loads it.

## [0.1.1] - 2026-09-28

### Fixed

- `gem install ruby_pptx` failed on Windows while generating ri
  documentation (`Errno::ENOENT` on a very long `.ri` path). RDoc 7 and
  earlier, which Ruby 3.3 to 4.0 bundle, cannot parse endless method
  definitions (`def x = ...`); each one ended its class early, so every later
  class was documented inside the one before it. On other platforms the
  install succeeded but the ri documentation was wrong. The library no longer
  uses endless methods, and CI checks that RDoc places every class correctly.
- `Presentation.open(path)` left a file handle open for every member of the
  package until garbage collection, so on Windows the `.pptx` stayed locked
  and could not be deleted or overwritten after opening it.

## [0.1.0] - 2026-09-28

The first release: a Ruby port of [python-pptx](https://github.com/scanny/python-pptx)
1.0.2, covering its whole public API, with an API redesigned for Ruby over an
object model kept close to the original.

### Added

**Everything python-pptx does.** Every public class and member of python-pptx
1.0.2 has a counterpart; `spec/ruby_pptx/api_completeness_spec.rb` checks this
against python-pptx itself and fails if a gap appears.

- Presentations: open, create from the default template, save; slide size;
  core properties.
- Slides, layouts and masters: add slides; find layouts by name and delete
  unused ones; backgrounds; speaker notes (`slide.notes = "…"`) and the notes
  master.
- Shapes: auto shapes with adjustment handles, text boxes, pictures, movies,
  connectors attached to shapes, groups, freeforms, tables with merged cells,
  and OLE objects shown as icons.
- Placeholders, with geometry inherited from layout and master, and
  `insert_picture` (cropped to fit), `insert_table` and `insert_chart`.
- Text: paragraphs, runs, fonts, alignment, spacing, autofit, hyperlinks and
  click actions, and `fit_text` to shrink text into its shape.
- Formatting: solid, patterned, gradient and background fills; line width,
  colour and dash style; shadows; theme and RGB colours with brightness.
- Pictures: PNG, JPEG, GIF, BMP, TIFF, WMF and EMF, read without Pillow;
  cropping; masking with a shape; reading the embedded image back.
- Charts: bar, column, line, pie, doughnut, area, radar, XY and bubble, in
  every python-pptx variant; grouped and date categories; replacing a chart's
  data; titles, legends, axes, tick marks and labels, gridlines, data labels
  per plot, series and point; markers; fonts; chart style; the exact chart
  type read back from an existing chart. The embedded workbook is written
  without XlsxWriter.

**Beyond python-pptx.**

- Slide sections.
- Paging a long table across as many slides as it needs.
- Combo charts, with a secondary axis.
- SVG pictures, with a raster fallback for older viewers.
- Defining a slide master, its theme and its layouts in code.
- `Pptx.build`, a declarative way to write a whole deck.

**A Ruby API.**

- Keyword arguments for position and size (`at:`, `size:`), which also accept
  `Pptx::Point` and `Pptx::Size` value objects with arithmetic.
- Explicit lengths (`Pptx.inches(1)`, `Pptx.pt(18)`), and `1.inch` through
  either the `Pptx::Lengths` refinement or the opt-in `ruby_pptx/core_ext`.
- Enumerable collections that index like arrays, ranges included.
- `case`/`in` pattern matching on shapes, slides, text, charts, lengths and
  colours.
- Predicates (`chart.legend?`, `frame.chart?`) and nil for "inherited" or
  "absent" throughout; reading a property never changes the document.

### Differences from python-pptx

Output matches python-pptx part for part, except where python-pptx is wrong
or where matching it cannot be done. Each case is listed, with its reason, in
[PORTING.md](PORTING.md):

- `fit_text` measures glyph outlines rather than rasterizing with Pillow, so
  it can choose a size one point different on text that only just fits. It
  also takes a `font_file:` rather than searching for one, since python-pptx's
  search works only on macOS and Windows.
- The chart's embedded workbook is written directly; its contents match, its
  bytes do not.
- python-pptx bugs this does not reproduce: `fit_text` raising `TypeError` on a
  word wider than the shape; reading the angle of a fresh gradient raising
  `TypeError`; an empty category label reading as `"None"`; an axis with no
  delete setting reading as hidden; EMF images stored as WMF.
- Getters that write to the document in python-pptx -- data-label flags, axis
  and chart titles -- do not here.

[0.1.1]: https://github.com/Largo/ruby_pptx/releases/tag/v0.1.1
[0.1.0]: https://github.com/Largo/ruby_pptx/releases/tag/v0.1.0
