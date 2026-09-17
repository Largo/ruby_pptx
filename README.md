# ruby_pptx

Create, read and update PowerPoint (`.pptx`) files from Ruby.

A port of [python-pptx](https://github.com/scanny/python-pptx) — the same
battle-tested OOXML object model underneath, with a public API redesigned for
Ruby. See [PORTING.md](PORTING.md) for the architecture and the milestone plan.

> **Status: in progress.** Presentations, slides, shapes, text, pictures,
> tables and charts work and are verified against python-pptx part-for-part.
> See [PORTING.md](PORTING.md) for what is and is not covered.

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

table = slide.shapes.add_table(2, 3, at: [Pptx.inches(1), Pptx.inches(3)],
                                     size: [Pptx.inches(8), Pptx.inches(2)]).table
table[0, 0].text = "Region"

data = Pptx::ChartData.new
data.categories = ["East", "West", "Midwest"]
data.add_series("Q1", [1.2, 2.0, 3.5])
slide.shapes.add_chart(:column_clustered, data,
                       at: [Pptx.inches(1), Pptx.inches(3)],
                       size: [Pptx.inches(8), Pptx.inches(4)])

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

Three things python-pptx does not do: **SVG pictures**, **slide sections**, and
**paging a long table across as many slides as it needs**.

```ruby
# PowerPoint wants a raster stand-in beside the vector, and this gem has no
# rasterizer, so you supply it.
slide.shapes.add_picture("logo.svg", at: [x, y], fallback: "logo.png")
```

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

require "pptx/core_ext"   # opt-in numeric sugar
1.inch == 72.points       #=> true
```

## Development

```bash
bundle install
bundle exec rspec
```

Specs compare the packages this gem writes against the ones python-pptx writes
for the same operation, part by part, on canonicalised XML. To run those:

```bash
pip install -r spec/oracle-requirements.txt
```

Without it, the 47 oracle-backed specs skip and the rest still run. CI sets
`REQUIRE_ORACLE=1`, which turns those skips into failures so a broken Python
environment cannot quietly reduce the suite to its unit tests.

## License

MIT — see [LICENSE](LICENSE).

This gem is a port of [python-pptx](https://github.com/scanny/python-pptx) by
Steve Canny, which is also MIT licensed, and it vendors four template files
from that project verbatim. See [NOTICE](NOTICE) for the full attribution and
the upstream licence text.
