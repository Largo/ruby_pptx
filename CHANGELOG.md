# Changelog

All notable changes to this gem are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html). Before 1.0 a minor
version may change the public API.

## [Unreleased]

### Fixed

Found by fuzzing the writer against the Open XML SDK's validator
(`.claude/skills/ooxml-fuzz`); each wrote XML the schema rejects:

- A new pattern colour (`fore_color`/`back_color` on a patterned fill) and a
  new gradient stop list (an outline gradient's `gradient_stops`) start from
  python-pptx's defaults instead of empty.
- A refused value no longer leaves a half-written element: `theme_color =
  :MIXED`, a `:MIXED` data-label position, a non-positive `major_unit` and
  the like raise and leave the XML as it was. `chart_style=` and `bullet=`
  keep their old value when refused.
- Control characters and U+FFFE/U+FFFF in names, core properties, notes,
  sections, chart text and theme names raise `ArgumentError` instead of
  producing a package PowerPoint cannot open. Run text still escapes them.
- `gap_width=` and `overlap=` raise `Pptx::Error` on plots other than bar
  and column (`bubble_scale=` on plots other than bubble), and a bubble
  series' `marker` raises, instead of writing elements those charts cannot
  hold.
- Data labels on a line chart with markers are written in schema order.
- `chart_style=` on a radar chart replaces its style rather than adding a
  second one.
- `bullet=` checks an autonumber scheme against the schema's list.
- `gradient_angle=` writes a whole turn (720, -360) as 0.
- A shape, chart, movie or OLE object with a negative size or one past
  2,147,483,647 EMU raises `ArgumentError` before anything is added, and so
  does a group that would grow that large. `width=` and `height=` raise
  `RangeError` for such a size, and a connector end that would land that far
  away (`begin_connect`, `end_x=`, ...) raises before anything moves.
- Chart axis ids are no longer negative; `c:axId` and `c:crossAx` are
  unsigned. The ids python-pptx writes negative are written without the sign.
- Radar chart series no longer end in a `c:smooth`, which radar series
  cannot hold.
- `major_unit=` and `minor_unit=` refuse NaN with `RangeError`.
- `brightness=` with any value but 0 raised `ArgumentError` ("wrong number
  of arguments") instead of adding the `a:lumMod`/`a:lumOff` adjustment. It
  now writes what python-pptx writes.

## [0.3.0] - 2026-11-17

### Added

- `Paragraph#bullet=` sets a character bullet (`"•"`), an autonumber scheme
  (`:arabicPeriod`), `:none`, or `nil` to inherit; `Paragraph#margin_left=`
  and `Paragraph#indent=` give it a hanging indent. python-pptx has neither.
- Layouts and masters take the drawing methods slides have: `add_shape`,
  `add_picture` (SVG with raster fallback included), `add_textbox`,
  `add_connector`, `add_freeform`/`build_freeform` and `add_group_shape` on
  `layout.shapes` and `master.shapes`. Every slide on the layout shows them,
  which is how a template carries its logo and bands. Charts, tables, movies
  and OLE objects stay slide-only. The `placeholders` views do not offer
  these methods, since a shape added there would not be listed.
- `SlideMaster#text_styles` reaches the master's `p:txStyles`: `title`,
  `body` and `other` are `ListStyle`s, and `level(n)` (0 to 8, like
  `Paragraph#level`) or `default` gives a `ParagraphStyle` with the same
  formatting methods as a paragraph -- alignment, spacing, `bullet`,
  `margin_left`, `indent` and `font`. Missing levels are created in schema
  order.
- `TextFrame#list_style`: the same `ListStyle` for a text body's
  `a:lstStyle`, so a layout placeholder can set what its slides inherit.
- `Paragraph#add_field(:slide_number)`, `(:date_time)` or any field type
  string such as `"datetime4"`.
- `TextFrame#text_direction=` with `:vertical_270` (bottom to top),
  `:vertical`, `:stacked` and the East Asian directions.
- `Font#caps=` with `:all`, `:small`, `:none` or `nil`.
- `Slides#add(layout, footers: true)` also copies the date, footer and
  slide-number placeholders, text included. PowerPoint only draws them where
  the slide carries them.
- Shapes drawn on a layout or master get `userDrawn="1"`, as PowerPoint
  writes it.
- `Chart#replace_data` rebinds scatter charts from an `XyChartData` and
  bubble charts from a `BubbleChartData`: it rewrites each series' `c:tx`,
  `c:xVal`, `c:yVal` and `c:bubbleSize` and keeps its formatting.
  ([#1](https://github.com/Largo/ruby_pptx/issues/1))

### Changed

- `Chart#replace_data` on a chart with more than one plot (a combo chart)
  raises `Pptx::Error` when the series count changes. Series are matched to
  plots in document order, so fewer series used to remove whole plots --
  leaving their secondary value axis behind -- and more series were all
  cloned into the last plot. Rebinding with the same number of series works
  as before. ([#2](https://github.com/Largo/ruby_pptx/issues/2))
- `Chart#replace_data` raises `Pptx::Error` when the data does not fit the
  chart: `ChartData` for a scatter or bubble chart, `XyChartData` for a
  category chart, or `XyChartData` without bubble sizes for a bubble chart.

### Fixed

- `Chart#replace_data(ChartData)` on a scatter chart wrote a `c:ser` with
  both `c:val` and the old `c:xVal`/`c:yVal`, which PowerPoint reports as
  unreadable content; it now raises instead. With an `XyChartData` it raised
  `NoMethodError`. ([#1](https://github.com/Largo/ruby_pptx/issues/1))
- Line series in a combo chart were drawn as smoothed curves: they carried
  no `c:smooth` of their own, which PowerPoint reads as smoothed regardless
  of the plot's setting. Each now writes `<c:smooth val="0"/>`, as a plain
  line chart does; `series.smooth = true` still curves one.
- The comment on `SlideLayout::LATENT_PLACEHOLDER_TYPES` said PowerPoint
  shows those placeholders from the layout; it does not.

## [0.2.0] - 2026-09-28

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

[0.3.0]: https://github.com/Largo/ruby_pptx/releases/tag/v0.3.0
[0.2.0]: https://github.com/Largo/ruby_pptx/releases/tag/v0.2.0
[0.1.1]: https://github.com/Largo/ruby_pptx/releases/tag/v0.1.1
[0.1.0]: https://github.com/Largo/ruby_pptx/releases/tag/v0.1.0
