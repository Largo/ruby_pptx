# ruby_pptx

Create, read and update PowerPoint (`.pptx`) files from Ruby.

A port of [python-pptx](https://github.com/scanny/python-pptx) — the same
battle-tested OOXML object model underneath, with a public API redesigned for
Ruby. See [PORTING.md](PORTING.md) for the architecture and the milestone plan.

> **Status:** every public class and member of python-pptx 1.0.2 has a
> counterpart here, which `spec/ruby_pptx/api_completeness_spec.rb` checks
> rather than asserts. Output is verified against python-pptx part for part.
> Where the two deliberately differ -- mostly python-pptx bugs this does not
> reproduce -- [PORTING.md](PORTING.md) lists each one.

```ruby
require "ruby_pptx"

prs = Pptx::Presentation.new_default          # or .open("deck.pptx")
slide = prs.slides.add(prs.slide_layouts["Title and Content"])

slide.shapes.title.text = "Quarterly Review"

body = slide.placeholders[1].text_frame
body.text = "Revenue up 12%\nCosts flat"
body.paragraphs.first.runs.first.font.tap do |font|
  font.bold = true
  font.size = Pptx.pt(24)
  font.color.rgb = Pptx::RGBColor["C0504D"]
end

box = slide.shapes.add_shape(:rounded_rectangle,
                             at: [Pptx.inches(1), Pptx.inches(5)],
                             size: [Pptx.inches(3), Pptx.inches(1)])
box.fill.solid
box.fill.fore_color.rgb = Pptx::RGBColor["1F497D"]
box.text_frame.text = "Next steps"

slide.shapes.add_picture("logo.png", at: [Pptx.inches(7), Pptx.inches(0.5)],
                         width: Pptx.inches(2))

table_frame = slide.shapes.add_table(2, 3, at: [Pptx.inches(1), Pptx.inches(3)],
                                     size: [Pptx.inches(8), Pptx.inches(2)])
table_frame.table[0, 0].text = "Region"

data = Pptx::ChartData.new
data.categories = ["East", "West", "Midwest"]
data.add_series("Q1", [1.2, 2.0, 3.5])
slide.shapes.add_chart(:column_clustered, data,
                       at: [Pptx.inches(1), Pptx.inches(3)],
                       size: [Pptx.inches(8), Pptx.inches(4)])

# Scatter and bubble charts take points rather than a value per category.
xy = Pptx::XyChartData.new
xy.add_series("Alpha", points: [[1, 10], [2, 20]])

box.hyperlink = "https://example.com"

# Connectors attach to shapes, and groups size themselves around their contents.
line = slide.shapes.add_connector(:straight, begin_at: [0, 0], end_at: [0, 0])
line.begin_connect(box, 3)
group = slide.shapes.add_group_shape([box, table_frame])

slide.shapes.add_movie("clip.mp4", at: [x, y], size: [w, h], content_type: "video/mp4")

prs.save("out.pptx")
```

Or build a whole deck declaratively:

```ruby
deck = Pptx.build do |d|
  d.slide_size = :widescreen

  d.slide("Title Slide") do |s|
    s.title = "Annual Report"
    s.subtitle = "Prepared in Ruby"
  end

  d.section("Detail") do
    d.slide("Blank") do |s|
      s.shape :rounded_rectangle, at: [Pptx.inches(1), Pptx.inches(1)],
                                  size: [Pptx.inches(3), Pptx.inches(1)],
                                  fill: "1F497D", text: "Next steps"
      s.chart :column_clustered, categories: %w[East West],
                                 series: { "Q1" => [1, 2] },
                                 at: [Pptx.inches(1), Pptx.inches(3)],
                                 size: [Pptx.inches(6), Pptx.inches(4)]
    end
  end
end

deck.save("out.pptx")
```

Five things python-pptx does not do: **SVG pictures**, **slide sections**,
**paging a long table across as many slides as it needs**, **combo charts**
with a secondary axis, and **defining a slide master in code**.

```ruby
data = Pptx::ChartData.new
data.categories = %w[Q1 Q2 Q3]
data.add_series("Revenue", [120, 135, 150])
data.add_series("Margin", [0.21, 0.24, 0.22])

slide.shapes.add_combo_chart(data, at: [x, y], size: [w, h]) do |combo|
  combo.plot :column_clustered, series: "Revenue"
  combo.plot :line, series: "Margin", secondary_axis: true
end
```

```ruby
# PowerPoint wants a raster stand-in beside the vector, and this gem has no
# rasterizer, so you supply it.
slide.shapes.add_picture("logo.svg", at: [x, y], fallback: "logo.png")
```

Text can be shrunk to fit the shape holding it. Measuring needs the actual
glyph outlines, so you pass the font file; `font_family` is what gets written
into the deck.

```ruby
box.text_frame.fit_text(font_file: "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                        font_family: "DejaVu Sans",
                        max_size: 28)
#=> 28   (the point size applied, never above max_size)
```

That turns word wrap on, autofit off, and applies the size to every run. Sizes
are measured from the font's own metrics rather than by rendering, so they can
differ from python-pptx's by a point on text that only just fits; PORTING.md
has the measured comparison.

A slide master can be built from scratch, rather than only edited in a
template. The master gets its own theme, so its colours and fonts are
independent of any other master in the deck.

```ruby
master = deck.slide_masters.add(name: "Corporate")
master.theme.colors.update(accent1: "1F497D", accent2: "C0504D")
master.theme.fonts.major = "Georgia"
master.theme.fonts.minor = "Verdana"

layout = master.slide_layouts.add("Title and Content", type: "obj") do |l|
  l.placeholders.add(:title, at: [Pptx.inches(0.5), Pptx.inches(0.3)],
                             size: [Pptx.inches(9), Pptx.inches(1.25)])
  l.placeholders.add(:body, idx: 1, at: [Pptx.inches(0.5), Pptx.inches(1.75)],
                                    size: [Pptx.inches(9), Pptx.inches(4.5)])
end

slide = deck.slides.add(layout)
slide.shapes.title.text = "Built from a hand-made master"
```

The master starts with the five placeholders PowerPoint expects — title, body,
date, footer and slide number — scaled to the deck's slide size. Pass
`placeholders: :none` for a bare one.

```ruby
deck.sections.add("Appendix", slides: deck.slides.to_a.last(2))

deck.slides.add_table_pages(rows,
                            layout: deck.slide_layouts["Blank"],
                            left: Pptx.inches(0.5), top: Pptx.inches(1),
                            width: Pptx.inches(9), height: Pptx.inches(5))
```

Lengths are explicit rather than bare numbers:

```ruby
Pptx.inches(1).emu        #=> 914400
Pptx.cm(2.54).pt          #=> 72.0
```

Numeric sugar is opt-in, and comes two ways. Prefer the refinement: it is
scoped to the file that asks for it, so it cannot surprise anything else
sharing the process.

```ruby
require "ruby_pptx/refinements"
using Pptx::Lengths               # this file only
1.inch == 72.points               #=> true

require "ruby_pptx/core_ext"      # or patch Numeric process-wide
```

`at:` and `size:` have always taken a two-element array, and still do. Passing
a `Point` or a `Size` instead costs nothing and buys named readers, arithmetic
and pattern matching:

```ruby
origin = Pptx.point(Pptx.inches(1), Pptx.inches(1))
box    = Pptx.size(Pptx.inches(3), Pptx.inches(1))

slide.shapes.add_shape(:rectangle, at: origin, size: box)
slide.shapes.add_shape(:rectangle, at: origin + [0, Pptx.inches(1.5)], size: box * 2)
```

Reading a deck back supports `case`/`in`. Enum-valued attributes read as their
symbolic name in a pattern, and only the keys a pattern asks for are computed:

```ruby
slide.shapes.each do |shape|
  case shape
  in {shape_type: :PICTURE, name:}                      then puts "picture #{name}"
  in {shape_type: :PLACEHOLDER, placeholder_format: {type: :TITLE}}
                                                        then puts shape.text_frame.text
  in {width:} if width > Pptx.inches(5)                 then puts "#{shape.name} is wide"
  else next
  end
end
```

Collections index like arrays — `slides[2]`, `slides[-1]`, `slides[1..3]`,
`slides[1, 2]` — and deconstruct into array patterns.

## Development

```bash
bundle install
bundle exec rspec
bundle exec rubocop
```

`.rubocop.yml` records where this codebase deliberately departs from the
default style — chiefly that the `oxml` layer keeps the OOXML schema's own
names (`CT_Shape`, `#cSld`, `accent1`) so the XML, the specification and the
Ruby read side by side.

Specs compare the packages this gem writes against the ones python-pptx writes
for the same operation, part by part, on canonicalised XML. To run those:

```bash
pip install -r spec/oracle-requirements.txt
```

Without it, the oracle-backed specs skip and the rest still run. CI sets
`REQUIRE_ORACLE=1`, which turns those skips into failures so a broken Python
environment cannot quietly reduce the suite to its unit tests. The same goes
for the `fit_text` specs, which measure with fonts from Debian's
`fonts-dejavu-core` and `fonts-urw-base35` packages.

A slide master has no oracle — python-pptx can read one but not create one —
so it is validated against the published ISO/IEC 29500-4 schemas instead.
Those are not redistributed here; point `OOXML_SCHEMAS` at a directory holding
`pml.xsd`, `dml-main.xsd` and the `shared-*.xsd` files they import:

```bash
OOXML_SCHEMAS=/path/to/schemas bundle exec rspec
```

## License

MIT — see [LICENSE](LICENSE).

This gem is a port of [python-pptx](https://github.com/scanny/python-pptx) by
Steve Canny, which is also MIT licensed, and it vendors nine template files
(the default deck, notes and theme templates, the video poster frame and the
OLE object icons) from that project verbatim, plus one derived from them. See
[NOTICE](NOTICE) for the full attribution and the upstream licence text.
