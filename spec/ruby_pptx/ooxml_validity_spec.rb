# frozen_string_literal: true

require "zip"

# Regressions found by fuzzing the writer and validating every deck with the
# Open XML SDK (.claude/skills/ooxml-fuzz). Each example is the minimized
# reproduction of one finding: a call sequence that used to write XML the
# schema rejects, or that raised but left half-written XML behind.
#
# Where python-pptx writes the same invalid XML, the example pins its output
# too (see "Upstream bugs this port does not reproduce" in PORTING.md), so an
# upstream fix shows up here.
RSpec.describe "OOXML validity of what the writer produces" do
  let(:deck) { Pptx::Presentation.new_default }
  let(:slide) { deck.slides.add(deck.slide_layouts[6]) }
  let(:box) { [Pptx.inches(1), Pptx.inches(1)] }
  let(:shape) { slide.shapes.add_shape(:rectangle, at: [0, 0], size: box) }

  def part_xml(name)
    blob = deck.to_blob
    blob = blob.string if blob.respond_to?(:string)
    Zip::File.open_buffer(StringIO.new(blob)).read(name).force_encoding("UTF-8")
  end

  def chart_data
    Pptx::ChartData.new.tap do |data|
      data.categories = %w[East West]
      data.add_series("Q1", [1, 2])
    end
  end

  def add_chart(type)
    slide.shapes.add_chart(type, chart_data, at: [0, 0], size: [Pptx.inches(4), Pptx.inches(3)]).chart
  end

  describe "pattern fills" do
    # An `a:fgClr` or `a:bgClr` must hold a colour. They start as python-pptx's
    # black and white, so asking for the back colour of a fresh pattern no
    # longer leaves an empty element.
    it "gives a new pattern colour slot a colour, as python-pptx does" do
      connector = slide.shapes.add_connector(:straight, begin_at: [0, 0], end_at: box)
      connector.line.fill.patterned
      connector.line.fill.fore_color
      connector.line.fill.back_color

      xml = part_xml("ppt/slides/slide1.xml")
      aggregate_failures do
        expect(xml).to include('<a:fgClr><a:srgbClr val="000000"/></a:fgClr>')
        expect(xml).to include('<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>')
      end
    end

    it "matches python-pptx for a patterned shape fill" do
      require_oracle!
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          from pptx.dml.color import RGBColor
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          shape = slide.shapes.add_shape(1, 0, 0, Inches(1), Inches(1))
          shape.fill.patterned()
          shape.fill.fore_color.rgb = RGBColor(0x1F, 0x49, 0x7D)
          shape.fill.back_color
          prs.save(out)
        PY
        ruby: lambda { |path|
          shape.fill.patterned
          shape.fill.fore_color.rgb = "1F497D"
          shape.fill.back_color
          deck.save(path)
        }
      )
    end
  end

  describe "gradients" do
    # python-pptx starts a new `a:gsLst` from two default stops; the port
    # missed that, so the stops of an outline gradient -- which starts empty
    # -- were an `a:gsLst` with no `a:gs`, which the schema rejects.
    it "gives an outline gradient's stops python-pptx's defaults" do
      require_oracle!
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          shape = slide.shapes.add_shape(1, 0, 0, Inches(1), Inches(1))
          shape.line.fill.gradient()
          shape.line.fill.gradient_stops
          prs.save(out)
        PY
        ruby: lambda { |path|
          shape.line.fill.gradient
          shape.line.fill.gradient_stops
          deck.save(path)
        }
      )
    end

    # 720 degrees used to normalize to 360 and be written as 21600000, which
    # ST_PositiveFixedAngle excludes; python-pptx does the same.
    it "writes a whole turn as angle zero" do
      shape.fill.gradient
      [720, -360, 360].each do |angle|
        shape.fill.gradient_angle = angle
        expect(shape.element.spPr.eg_fillProperties.lin.get("ang")).to eq("0")
      end
    end
  end

  describe "colours" do
    it "refuses a theme colour with no XML value and keeps the colour it had" do
      shape.fill.solid
      shape.fill.fore_color.rgb = "112233"
      aggregate_failures do
        expect { shape.fill.fore_color.theme_color = :MIXED }.to raise_error(ArgumentError)
        expect(shape.fill.fore_color.rgb.to_s).to eq("112233")
        expect(part_xml("ppt/slides/slide1.xml")).not_to match(/<a:schemeClr(?![^>]*val=)/)
      end
    end

    # Any non-zero brightness used to raise "wrong number of arguments": the
    # generated get_or_add_lumMod called the one-argument add_lumMod.
    it "matches python-pptx for a lightened theme colour and a darkened RGB one" do
      require_oracle!
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          from pptx.dml.color import RGBColor
          from pptx.enum.dml import MSO_THEME_COLOR
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          shape = slide.shapes.add_shape(1, 0, 0, Inches(1), Inches(1))
          shape.fill.solid()
          shape.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
          shape.fill.fore_color.brightness = 0.4
          shape.line.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
          shape.line.color.brightness = -0.25
          prs.save(out)
        PY
        ruby: lambda { |path|
          shape.fill.solid
          shape.fill.fore_color.theme_color = :ACCENT_1
          shape.fill.fore_color.brightness = 0.4
          shape.line.color.rgb = "1F497D"
          shape.line.color.brightness = -0.25
          deck.save(path)
        }
      )
    end
  end

  describe "characters XML cannot carry" do
    it "still escapes control characters in run text, as python-pptx does" do
      run = slide.shapes.add_textbox(at: [0, 0], size: box).text_frame.paragraphs[0].add_run
      run.text = "a\u0001b"
      expect(part_xml("ppt/slides/slide1.xml")).to include("a_x0001_b")
    end

    # python-pptx raises ValueError for these through lxml. Here they used to
    # reach the package and make the part unparseable.
    {
      "a shape name" => ->(t) { t.shape.name = "a\vb" },
      "a master name" => ->(t) { t.deck.slide_masters.add(name: "a\vb") },
      "a core property" => ->(t) { t.deck.core_properties.title = "a\u0000b" },
      "notes" => ->(t) { t.slide.notes = "a\u{FFFE}b" },
      "a chart category" => lambda { |t|
        data = Pptx::ChartData.new
        data.categories = ["a\vb"]
        data.add_series("S", [1])
        t.slide.shapes.add_chart(:pie, data, at: [0, 0], size: t.box)
      }
    }.each do |where, call|
      it "refuses one in #{where}" do
        expect { call.call(self) }.to raise_error(ArgumentError, /not XML compatible/)
      end
    end

    it "refuses one in a section name without leaving an empty section list" do
      aggregate_failures do
        expect { deck.sections.add("a\vb") }.to raise_error(ArgumentError)
        expect(deck.element.section_list).to be_nil
      end
    end
  end

  describe "shape sizes" do
    it "refuses a negative size before adding the shape" do
      shapes = slide.shapes
      aggregate_failures do
        expect { shapes.add_shape(:rectangle, at: [0, 0], size: [Pptx.emu(-1), Pptx.emu(1)]) }
          .to raise_error(ArgumentError, /width and height/)
        expect(shapes.size).to eq(0)
      end
    end

    it "refuses a width or height past what Office reads and keeps the old one" do
      aggregate_failures do
        expect { shape.width = 2**31 }.to raise_error(RangeError)
        expect { shape.height = -1 }.to raise_error(RangeError)
        expect([shape.width, shape.height]).to eq(box)
      end
    end

    it "refuses to connect to a point too far away and leaves the connector as it was" do
      wide = slide.shapes.add_textbox(at: [Pptx.inches(8), 0], size: [(2**31) - 1, Pptx.inches(1)])
      connector = slide.shapes.add_connector(:STRAIGHT, begin_at: [Pptx.inches(7), 0], end_at: [0, 0])
      before = connector.element.to_xml
      aggregate_failures do
        expect { connector.begin_connect(wide, 3) }.to raise_error(ArgumentError)
        expect(connector.element.to_xml).to eq(before)
        expect(connector).not_to be_begin_connected
      end
    end

    it "refuses a chart too large for Office before adding its part" do
      parts_before = slide.part.rels.size
      huge = [Pptx.emu(2**31), Pptx.emu(1)]
      aggregate_failures do
        expect { slide.shapes.add_chart(:pie, chart_data, at: [0, 0], size: huge) }
          .to raise_error(ArgumentError)
        expect(slide.part.rels.size).to eq(parts_before)
      end
    end
  end

  describe "groups" do
    it "refuses a group too wide for Office and leaves its shapes on the slide" do
      near = slide.shapes.add_shape(:rectangle, at: [0, 0], size: box)
      far = slide.shapes.add_shape(:rectangle, at: [Pptx.emu((2**31) - 1), 0], size: box)
      aggregate_failures do
        expect { slide.shapes.add_group_shape([near, far]) }.to raise_error(ArgumentError)
        expect(slide.shapes.map(&:shape_type).map(&:name)).to eq(%i[AUTO_SHAPE AUTO_SHAPE])
      end
    end
  end

  describe "shapes added to a group" do
    it "refuses one that would stretch the group too far and leaves the group as it was" do
      inner = slide.shapes.add_shape(:rectangle, at: [0, Pptx.emu(102_412_800)],
                                                 size: [box[0], Pptx.emu((2**31) - 1)])
      group = slide.shapes.add_group_shape([inner])
      before = [group.left, group.top, group.width, group.height]
      aggregate_failures do
        expect { group.shapes.add_shape(:oval, at: [0, 0], size: box) }.to raise_error(ArgumentError)
        expect(group.shapes.size).to eq(1)
        expect([group.left, group.top, group.width, group.height]).to eq(before)
      end
    end
  end

  describe "bullets" do
    it "refuses a numbering scheme the schema does not list and keeps the old bullet" do
      paragraph = slide.shapes.add_textbox(at: [0, 0], size: box).text_frame.paragraphs[0]
      paragraph.bullet = :arabicPeriod
      aggregate_failures do
        expect { paragraph.bullet = :arabicperiod }.to raise_error(ArgumentError)
        expect(paragraph.bullet).to eq(:arabicPeriod)
      end
    end
  end

  describe "charts" do
    it "allows gap width and overlap on bar plots only" do
      line = add_chart(:line).plots.first
      bar = add_chart(:column_clustered).plots.first
      aggregate_failures do
        expect { line.gap_width = 80 }.to raise_error(Pptx::Error, /no gap width/)
        expect { line.overlap = 20 }.to raise_error(Pptx::Error, /no overlap/)
        expect { line.bubble_scale = 50 }.to raise_error(Pptx::Error, /no bubble scale/)
        bar.gap_width = 80
        expect(bar.gap_width).to eq(80)
      end
    end

    # python-pptx's BubbleSeries inherits XySeries#marker and writes a
    # c:marker that CT_BubbleSer does not allow.
    it "has no marker on a bubble series" do
      data = Pptx::BubbleChartData.new
      data.add_series("S", points: [[1, 2, 3]])
      chart = slide.shapes.add_chart(:bubble, data, at: [0, 0], size: box).chart
      series = chart.series.first
      aggregate_failures do
        expect { series.marker }.to raise_error(Pptx::Error, /no marker/)
        expect(part_xml("ppt/charts/chart1.xml")).not_to include("<c:marker")
      end
    end

    # `c:marker` and `c:smooth` follow `c:dLbls` in a line chart; labels used
    # to be appended after them.
    it "puts a line chart's data labels before its marker and smooth flags" do
      add_chart(:line_markers_stacked).plots.first.data_labels = true
      plot = part_xml("ppt/charts/chart1.xml")[%r{<c:lineChart>.*</c:lineChart>}m]
      expect(plot.index("<c:dLbls>")).to be < plot.index("<c:marker ")
    end

    # The radar template, like PowerPoint, keeps its style in an
    # mc:AlternateContent. python-pptx adds a second, bare c:style after it.
    it "replaces a radar chart's style rather than adding a second one" do
      add_chart(:radar).chart_style = 7
      xml = part_xml("ppt/charts/chart1.xml")
      aggregate_failures do
        expect(xml.scan("<c:style ").size).to eq(1)
        expect(xml).not_to include("AlternateContent")
      end
    end

    # python-pptx's templates carry negative axis ids, which c:axId and
    # c:crossAx (xsd:unsignedInt) do not allow; Office reads at most 2**31 - 1.
    it "writes every axis id as an unsigned int Office reads" do
      xy = Pptx::XyChartData.new.tap { |data| data.add_series("S", points: [[1, 2]]) }
      bubbles = Pptx::BubbleChartData.new.tap { |data| data.add_series("S", points: [[1, 2, 3]]) }
      %i[column_clustered line area radar].each { |type| add_chart(type) }
      slide.shapes.add_chart(:xy_scatter, xy, at: [0, 0], size: box)
      slide.shapes.add_chart(:bubble, bubbles, at: [0, 0], size: box)
      combo = Pptx::ChartData.new.tap do |data|
        data.categories = %w[East West]
        data.add_series("A", [1, 2])
        data.add_series("B", [3, 4])
      end
      slide.shapes.add_combo_chart(combo, at: [0, 0], size: box) do |c|
        c.plot :column_clustered, series: "A"
        c.plot :line, series: "B", secondary_axis: true
      end
      ids = (1..7).flat_map do |n|
        part_xml("ppt/charts/chart#{n}.xml").scan(/<c:(?:axId|crossAx) val="([^"]*)"/).flatten
      end
      expect(ids).to all(match(/\A\d+\z/).and(satisfy { |id| id.to_i <= (2**31) - 1 }))
    end

    # CT_RadarSer has no c:smooth; python-pptx's radar template copies the
    # line chart's series and ends each with one.
    it "writes no c:smooth in a radar series" do
      %i[radar radar_markers radar_filled].each { |type| add_chart(type) }
      xml = (1..3).map { |n| part_xml("ppt/charts/chart#{n}.xml") }.join
      aggregate_failures do
        expect(xml.scan("<c:ser>").size).to eq(3)
        expect(xml).not_to include("<c:smooth")
      end
    end

    # The series' c:smooth, not scatterStyle, decides whether the line curves.
    it "smooths the series of a smooth scatter chart and only those" do
      xy = Pptx::XyChartData.new.tap do |data|
        data.add_series("A", points: [[1, 2], [2, 5]])
        data.add_series("B", points: [[1, 1], [2, 2]])
      end
      types = %i[xy_scatter_smooth xy_scatter_smooth_no_markers
                 xy_scatter_lines xy_scatter_lines_no_markers xy_scatter]
      types.each { |type| slide.shapes.add_chart(type, xy, at: [0, 0], size: box) }
      smooth = (1..5).map { |n| part_xml("ppt/charts/chart#{n}.xml").scan(/<c:smooth val="(\d)"/).flatten }
      expect(smooth).to eq([%w[1 1], %w[1 1], %w[0 0], %w[0 0], %w[0 0]])
    end

    it "keeps the style it had when given one out of range" do
      chart = add_chart(:column_clustered)
      chart.chart_style = 7
      aggregate_failures do
        expect { chart.chart_style = 99 }.to raise_error(RangeError)
        expect(chart.chart_style).to eq(7)
      end
    end

    # Each of these used to add the element and then fail on its value,
    # leaving it without the attribute the schema requires.
    it "leaves nothing behind when a value is refused" do
      chart = add_chart(:column_clustered)
      labels = chart.plots.first.data_labels
      aggregate_failures do
        expect { labels.position = :MIXED }.to raise_error(ArgumentError)
        expect { chart.value_axis.major_unit = -1 }.to raise_error(RangeError)
        expect { chart.value_axis.minor_unit = Float::NAN }.to raise_error(RangeError)
        xml = part_xml("ppt/charts/chart1.xml")
        expect(xml).not_to include("<c:dLblPos")
        expect(xml).not_to include("<c:majorUnit")
        expect(xml).not_to include("<c:minorUnit")
      end
    end
  end

  # The divergences above, from python-pptx's side: each of these writes XML
  # the schema rejects, or draws other than asked. Asserted so that an
  # upstream fix shows up here.
  describe "what python-pptx 1.0.2 writes for the same calls" do
    before { require_oracle! }

    def python_xml(body, part)
      script = <<~PY
        from pptx.util import Inches, Emu
        from pptx.chart.data import CategoryChartData, BubbleChartData, XyChartData
        from pptx.enum.chart import XL_CHART_TYPE
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        #{body.gsub("\n", "\n")}
        prs.save(out)
      PY
      python_pptx_manifest(script).fetch("parts").fetch(part).fetch("c14n")
    end

    it "writes a whole-turn gradient angle as 21600000" do
      xml = python_xml(<<~PY.strip, "ppt/slides/slide1.xml")
        shape = slide.shapes.add_shape(1, 0, 0, Inches(1), Inches(1))
        shape.fill.gradient()
        shape.fill.gradient_angle = 720
      PY
      expect(xml).to include('ang="21600000"')
    end

    it "writes a negative shape size" do
      xml = python_xml("slide.shapes.add_shape(1, 0, 0, Emu(-1), Inches(1))", "ppt/slides/slide1.xml")
      expect(xml).to include('cx="-1"')
    end

    it "adds a second c:style to a radar chart" do
      xml = python_xml(<<~PY.strip, "ppt/charts/chart1.xml")
        cd = CategoryChartData()
        cd.categories = ["East", "West"]
        cd.add_series("Q1", (1, 2))
        chart = slide.shapes.add_chart(XL_CHART_TYPE.RADAR, 0, 0, Inches(4), Inches(3), cd).chart
        chart.chart_style = 7
      PY
      expect(xml.scan("<c:style ").size).to eq(2)
    end

    it "writes negative axis ids and a c:smooth in each radar series" do
      xml = python_xml(<<~PY.strip, "ppt/charts/chart1.xml")
        cd = CategoryChartData()
        cd.categories = ["East", "West"]
        cd.add_series("Q1", (1, 2))
        slide.shapes.add_chart(XL_CHART_TYPE.RADAR, 0, 0, Inches(4), Inches(3), cd)
      PY
      aggregate_failures do
        expect(xml).to include('<c:axId val="-2112772216">')
        expect(xml[%r{<c:ser>.*</c:ser>}m]).to include('<c:smooth val="0">')
      end
    end

    it "writes c:smooth off in the series of a smooth scatter chart" do
      xml = python_xml(<<~PY.strip, "ppt/charts/chart1.xml")
        cd = XyChartData()
        cd.add_series("S").add_data_point(1, 2)
        slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_SMOOTH, 0, 0, Inches(4), Inches(3), cd)
      PY
      aggregate_failures do
        expect(xml).to include('<c:scatterStyle val="smoothMarker">')
        expect(xml[%r{<c:ser>.*</c:ser>}m]).to include('<c:smooth val="0">')
      end
    end

    it "writes a NaN axis unit" do
      xml = python_xml(<<~PY.strip, "ppt/charts/chart1.xml")
        cd = CategoryChartData()
        cd.categories = ["East", "West"]
        cd.add_series("Q1", (1, 2))
        chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 0, Inches(4), Inches(3), cd).chart
        chart.value_axis.major_unit = float("nan")
      PY
      expect(xml).to include('<c:majorUnit val="nan">')
    end

    it "writes a marker into a bubble series" do
      xml = python_xml(<<~PY.strip, "ppt/charts/chart1.xml")
        cd = BubbleChartData()
        cd.add_series("S").add_data_point(1, 2, 3)
        chart = slide.shapes.add_chart(XL_CHART_TYPE.BUBBLE, 0, 0, Inches(4), Inches(3), cd).chart
        chart.plots[0].series[0].marker.size = 10
      PY
      expect(xml).to include("<c:marker>")
    end
  end
end
