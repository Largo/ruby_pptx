# frozen_string_literal: true

require "json"

# Read parity: python-pptx builds a chart, then the same properties are read
# by both libraries and compared.
#
# The package differentials cannot see a wrong *read*. An absent c:varyColors
# was read as false here and as true by python-pptx; no file differed, so no
# differential noticed. This is the oracle for reads.
RSpec.describe "reading charts the way python-pptx does" do
  before { require_oracle! }

  READ_ORACLE = File.expand_path("../../tools/chart_read_oracle.py", __dir__)

  # python expression => how to read the same thing here. Every row is the
  # same shape of lambda, so the table reads down; symbol-procs for the few
  # one-method rows would break that.
  # rubocop:disable-next Style/SymbolProc
  READ_PROBES = {
    "chart.chart_type" => ->(c) { c.chart_type },
    "chart.chart_style" => ->(c) { c.chart_style },
    "chart.has_legend" => ->(c) { c.legend? },
    "chart.has_title" => ->(c) { c.title? },
    "chart.legend.position" => ->(c) { c.legend.position },
    "chart.legend.include_in_layout" => ->(c) { c.legend.include_in_layout? },
    "chart.legend.horz_offset" => ->(c) { c.legend.horz_offset },
    "chart.legend.font.size" => ->(c) { c.legend.font.size&.emu },
    "chart.font.size" => ->(c) { c.font.size&.emu },
    "chart.font.bold" => ->(c) { c.font.bold },
    "chart.plots[0].vary_by_categories" => ->(c) { c.plots[0].vary_by_categories? },
    "chart.plots[0].has_data_labels" => ->(c) { c.plots[0].data_labels? },
    "len(chart.plots[0].categories)" => ->(c) { c.plots[0].categories.size },
    "list(chart.plots[0].categories)" => ->(c) { c.plots[0].categories.map(&:label) },
    "chart.plots[0].categories.depth" => ->(c) { c.plots[0].categories.depth },
    "chart.plots[0].categories.flattened_labels" => ->(c) { c.plots[0].categories.flattened_labels },
    "[s.name for s in chart.series]" => ->(c) { c.series.map(&:name) },
    "[list(s.values) for s in chart.series]" => ->(c) { c.series.map(&:values) },
    "[s.index for s in chart.series]" => ->(c) { c.series.map(&:index) },
    "[len(s.points) for s in chart.series]" => ->(c) { c.series.map { |s| s.points.size } },
    "chart.category_axis.major_tick_mark" => ->(c) { c.category_axis.major_tick_mark },
    "chart.category_axis.minor_tick_mark" => ->(c) { c.category_axis.minor_tick_mark },
    "chart.category_axis.reverse_order" => ->(c) { c.category_axis.reverse_order? },
    "chart.category_axis.tick_label_position" => ->(c) { c.category_axis.tick_label_position },
    "chart.category_axis.has_major_gridlines" => ->(c) { c.category_axis.major_gridlines? },
    "chart.category_axis.has_title" => ->(c) { c.category_axis.title? },
    "chart.category_axis.tick_labels.offset" => ->(c) { c.category_axis.tick_labels.offset },
    "chart.value_axis.maximum_scale" => ->(c) { c.value_axis.maximum_scale },
    "chart.value_axis.minimum_scale" => ->(c) { c.value_axis.minimum_scale },
    "chart.value_axis.major_unit" => ->(c) { c.value_axis.major_unit },
    "chart.value_axis.minor_unit" => ->(c) { c.value_axis.minor_unit },
    "chart.value_axis.crosses" => ->(c) { c.value_axis.crosses },
    "chart.value_axis.crosses_at" => ->(c) { c.value_axis.crosses_at },
    "chart.value_axis.has_major_gridlines" => ->(c) { c.value_axis.major_gridlines? },
    "chart.value_axis.tick_labels.number_format" => ->(c) { c.value_axis.tick_labels.number_format },
    "chart.value_axis.tick_labels.number_format_is_linked" =>
      ->(c) { c.value_axis.tick_labels.number_format_linked? },
    "chart.value_axis.tick_labels.font.size" => ->(c) { c.value_axis.tick_labels.font.size&.emu },
    "chart.plots[0].data_labels.number_format" => ->(c) { c.plots[0].data_labels.number_format },
    "chart.plots[0].data_labels.number_format_is_linked" =>
      ->(c) { c.plots[0].data_labels.number_format_linked? },
    "chart.plots[0].data_labels.position" => ->(c) { c.plots[0].data_labels.position },
    "chart.plots[0].data_labels.show_value" => ->(c) { c.plots[0].data_labels.show_value? },
    "chart.plots[0].gap_width" => ->(c) { c.plots[0].gap_width },
    "chart.plots[0].overlap" => ->(c) { c.plots[0].overlap },
    "chart.plots[0].bubble_scale" => ->(c) { c.plots[0].bubble_scale },
    "chart.series[0].smooth" => ->(c) { c.series[0].smooth? },
    "chart.series[0].invert_if_negative" => ->(c) { c.series[0].invert_if_negative? },
    "chart.series[0].marker.style" => ->(c) { c.series[0].marker.style },
    "chart.series[0].marker.size" => ->(c) { c.series[0].marker.size },
    "chart.series[0].points[1].marker.style" => ->(c) { c.series[0].points[1].marker.style },
    "chart.series[0].points[1].data_label.position" => ->(c) { c.series[0].points[1].data_label.position },
    "chart.series[0].points[1].data_label.has_text_frame" =>
      ->(c) { c.series[0].points[1].data_label.text_frame? },
    "chart.series[0].points[1].data_label.text_frame.text" =>
      ->(c) { c.series[0].points[1].data_label.text_frame.text },
    "chart.series[0].data_labels.show_value" => ->(c) { c.series[0].data_labels.show_value? },
    "chart.value_axis.axis_title.text_frame.text" => ->(c) { c.value_axis.title&.text },
    "chart.chart_title.text_frame.text" => ->(c) { c.title&.text },
    "chart.category_axis.visible" => ->(c) { c.category_axis.visible? },
    "chart.value_axis.visible" => ->(c) { c.value_axis.visible? },
    "chart.category_axis.minimum_scale" => ->(c) { c.category_axis.minimum_scale },
    "[s.smooth for s in chart.series]" => ->(c) { c.series.map(&:smooth?) }
  }.freeze

  # Enum members read as names, as the oracle prints them.
  def normalise(value)
    case value
    when Pptx::Enum::Member then value.name.to_s
    when Pptx::ChartCategory then value.label
    when Array then value.map { |v| normalise(v) }
    else value
    end
  end

  def python_reads(path, slide, expressions)
    out, err, status = Open3.capture3("python3", READ_ORACLE, path, slide.to_s, JSON.dump(expressions))
    raise "read oracle failed: #{err}" unless status.success?

    JSON.parse(out)
  end

  # Builds every fixture chart in one python-pptx session, one per slide.
  def self.build_fixtures(path)
    script = <<~PY
      import sys, pptx
      from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData
      from pptx.enum.chart import (XL_CHART_TYPE as XL, XL_LEGEND_POSITION, XL_TICK_MARK,
          XL_TICK_LABEL_POSITION, XL_AXIS_CROSSES, XL_DATA_LABEL_POSITION, XL_MARKER_STYLE)
      from pptx.util import Inches, Pt
      prs = pptx.Presentation()

      def cat_data(values=((1.5, -2, 3), (4, 5, 6.25))):
          cd = CategoryChartData()
          cd.categories = ["East", "West", "Mid"]
          for i, v in enumerate(values):
              cd.add_series("S%d" % i, v)
          return cd

      def add(chart_type, data):
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          return slide.shapes.add_chart(chart_type, 0, 0, Inches(6), Inches(4), data).chart

      # 0: a fresh chart, nothing set -- every default
      add(XL.COLUMN_CLUSTERED, cat_data())

      # 1: a chart with nearly everything set
      c = add(XL.BAR_STACKED, cat_data())
      c.chart_style = 17
      c.has_legend = True
      c.legend.position = XL_LEGEND_POSITION.BOTTOM
      c.legend.include_in_layout = False
      c.legend.horz_offset = 0.2
      c.legend.font.size = Pt(9)
      c.font.size = Pt(11)
      c.font.bold = True
      c.chart_title.text_frame.text = "Revenue"
      ca, va = c.category_axis, c.value_axis
      ca.major_tick_mark = XL_TICK_MARK.OUTSIDE
      ca.minor_tick_mark = XL_TICK_MARK.INSIDE
      ca.reverse_order = True
      ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
      ca.has_major_gridlines = True
      ca.tick_labels.offset = 250
      ca.has_title = True
      va.maximum_scale = 12.5
      va.minimum_scale = -3
      va.major_unit = 2.5
      va.minor_unit = 0.5
      va.crosses = XL_AXIS_CROSSES.MAXIMUM
      va.has_major_gridlines = False
      va.tick_labels.number_format = "0.0%"
      va.tick_labels.font.size = Pt(8)
      va.axis_title.text_frame.text = "Units"
      p = c.plots[0]
      p.gap_width = 80
      p.overlap = 100
      p.vary_by_categories = False
      p.has_data_labels = True
      p.data_labels.number_format = "#,##0"
      p.data_labels.position = XL_DATA_LABEL_POSITION.INSIDE_END
      c.series[0].invert_if_negative = False

      # 2: a line chart, with markers, smoothing and per-point settings
      c = add(XL.LINE_MARKERS, cat_data())
      s = c.series[0]
      s.smooth = False
      s.marker.style = XL_MARKER_STYLE.DIAMOND
      s.marker.size = 11
      pt = s.points[1]
      pt.marker.style = XL_MARKER_STYLE.SQUARE
      pt.data_label.position = XL_DATA_LABEL_POSITION.ABOVE
      pt.data_label.text_frame.text = "peak"
      s.data_labels.show_value = True
      c.value_axis.crosses_at = 1.5

      # 3: an XY scatter, which has two value axes
      xy = XyChartData()
      s = xy.add_series("Alpha")
      for x, y in ((1, 10), (2, 20), (3, 15.5)):
          s.add_data_point(x, y)
      add(XL.XY_SCATTER_SMOOTH_NO_MARKERS, xy)

      # 4: a bubble chart
      bd = BubbleChartData()
      s = bd.add_series("B")
      for x, y, z in ((1, 10, 5), (2, 20, 7)):
          s.add_data_point(x, y, z)
      c = add(XL.BUBBLE, bd)
      c.plots[0].bubble_scale = 150

      # 5..10: chart types the inspector has to tell apart
      for t in (XL.PIE_EXPLODED, XL.DOUGHNUT, XL.RADAR_FILLED, XL.AREA_STACKED_100,
                XL.LINE_STACKED, XL.XY_SCATTER_LINES):
          data = xy if t == XL.XY_SCATTER_LINES else cat_data(((1, 2, 3),))
          add(t, data)

      # 11: categories in two levels
      cd = CategoryChartData()
      east = cd.categories.add_category("East")
      east.add_sub_category("North")
      east.add_sub_category("South")
      west = cd.categories.add_category("West")
      west.add_sub_category("Coast")
      cd.add_series("S", (1, 2, 3))
      add(XL.COLUMN_CLUSTERED, cd)

      # The fixtures above almost never leave an optional element out --
      # python-pptx's templates write them -- so the defaults for an absent
      # element were never read. These force it.

      # 12: settings restored to their defaults, which removes the elements
      c = add(XL.COLUMN_CLUSTERED, cat_data())
      c.has_legend = True
      ca = c.category_axis
      ca.major_tick_mark = XL_TICK_MARK.CROSS
      ca.minor_tick_mark = XL_TICK_MARK.CROSS
      ca.tick_labels.offset = 100
      c.value_axis.visible = False

      # 13: smoothing bare and absent, and series drawn out of document order
      c = add(XL.LINE, cat_data())
      s0, s1 = c.series[0], c.series[1]
      s0.smooth = True
      s1._element.remove(s1._element.smooth)
      o0, o1 = s0._element.order, s1._element.order
      o0.val, o1.val = 1, 0

      # 14: an XY chart whose two value axes differ, so which is which shows
      c = add(XL.XY_SCATTER, xy)
      c.value_axis.maximum_scale = 50
      c.category_axis.minimum_scale = -5

      # 15: an empty category, which has no c:pt but still counts
      cd = CategoryChartData()
      cd.categories = ["East", None, "Mid"]
      cd.add_series("S", (1, 2, 3))
      add(XL.COLUMN_CLUSTERED, cd)

      # 16: a category with no c:pt at all, counted only by c:ptCount
      cd = CategoryChartData()
      cd.categories = ["East", "West", "Mid"]
      cd.add_series("S", (1, 2, 3))
      c = add(XL.COLUMN_CLUSTERED, cd)
      for pt in c._chartSpace.xpath(".//c:ser/c:cat//c:pt[@idx='1']"):
          pt.getparent().remove(pt)

      prs.save(sys.argv[1])
    PY
    _, err, status = Open3.capture3("python3", "-c", script, path)
    raise "fixture build failed: #{err}" unless status.success?
  end

  # Which probes each slide is read with. A probe a chart cannot answer --
  # a category axis on a pie -- is left out rather than compared.
  SERIES_PROBES = ["[s.name for s in chart.series]", "[list(s.values) for s in chart.series]",
                   "[s.index for s in chart.series]", "[len(s.points) for s in chart.series]"].freeze
  COMMON_PROBES = READ_PROBES.keys.grep(/\Achart\.(chart_type|chart_style|has_|font)/) + SERIES_PROBES
  # python-pptx's axis_title creates a title in order to read it, so on an
  # axis with none it answers "" and has changed the file; this answers nil
  # and has not. That probe is only compared where a title exists.
  AXIS_PROBES = READ_PROBES.keys.select { |k| k.include?("_axis.") && !k.include?("axis_title") }
  AXIS_TITLE_PROBES = ["chart.value_axis.axis_title.text_frame.text"].freeze
  CATEGORY_PROBES = READ_PROBES.keys.grep(/categories/)
  LEGEND_PROBES = READ_PROBES.keys.grep(/\Achart\.legend/)
  TITLE_PROBES = READ_PROBES.keys.grep(/\Achart\.chart_title/)
  PLOT_PROBES = READ_PROBES.keys.grep(/plots\[0\]\.(vary|has_data|gap|overlap)/)
  DATA_LABEL_PROBES = READ_PROBES.keys.grep(/plots\[0\]\.data_labels/)
  BAR_SERIES_PROBES = READ_PROBES.keys.grep(/series\[0\]\.invert/)
  LINE_SERIES_PROBES = READ_PROBES.keys.grep(/series\[0\]\.(smooth|marker|points|data_labels)/)
  RESTORED_DEFAULT_PROBES = ["chart.category_axis.major_tick_mark", "chart.category_axis.minor_tick_mark",
                             "chart.category_axis.tick_labels.offset", "chart.value_axis.visible",
                             "chart.category_axis.visible"].freeze
  XY_AXIS_PROBES = ["chart.value_axis.maximum_scale", "chart.category_axis.minimum_scale",
                    "chart.value_axis.minimum_scale"].freeze

  PROBES_BY_SLIDE = {
    0 => COMMON_PROBES + AXIS_PROBES + CATEGORY_PROBES + PLOT_PROBES + BAR_SERIES_PROBES,
    1 => COMMON_PROBES + LEGEND_PROBES + TITLE_PROBES + AXIS_PROBES + AXIS_TITLE_PROBES +
         PLOT_PROBES + DATA_LABEL_PROBES + BAR_SERIES_PROBES,
    2 => COMMON_PROBES + LINE_SERIES_PROBES + ["chart.value_axis.crosses", "chart.value_axis.crosses_at"],
    3 => COMMON_PROBES + ["chart.value_axis.maximum_scale", "chart.value_axis.crosses"],
    4 => COMMON_PROBES + ["chart.plots[0].bubble_scale"],
    # 5..10 exist for chart_type, which COMMON_PROBES covers.
    **(5..10).to_h { |slide| [slide, COMMON_PROBES] },
    11 => COMMON_PROBES + CATEGORY_PROBES,
    12 => COMMON_PROBES + LEGEND_PROBES + RESTORED_DEFAULT_PROBES,
    13 => COMMON_PROBES + ["[s.smooth for s in chart.series]"],
    14 => COMMON_PROBES + XY_AXIS_PROBES,
    15 => COMMON_PROBES + CATEGORY_PROBES,
    16 => COMMON_PROBES + CATEGORY_PROBES
  }.freeze

  # Reads where python-pptx is wrong and this is not. Both values are
  # asserted, so a fix upstream shows up here rather than going unnoticed.
  KNOWN_READ_DIVERGENCES = {
    # An empty label is <c:v/>, whose text lxml gives as None; python-pptx
    # then builds its str-subclass Category from it and gets "None".
    [15, "list(chart.plots[0].categories)"] =>
      { python: %w[East None Mid], ruby: ["East", "", "Mid"] },
    [15, "chart.plots[0].categories.flattened_labels"] =>
      { python: [["East"], ["None"], ["Mid"]], ruby: [["East"], [""], ["Mid"]] }
  }.freeze

  # Built once for the whole file.
  before(:context) do
    require_oracle!
    @fixture = Tempfile.new(["read-parity", ".pptx"])
    @fixture.close
    self.class.build_fixtures(@fixture.path)
    @deck = Pptx::Presentation.open(@fixture.path)
  end

  after(:context) { @fixture&.unlink }

  PROBES_BY_SLIDE.each do |slide, probes|
    it "reads slide #{slide} the way python-pptx does" do
      chart = @deck.slides[slide].shapes.find(&:chart?).chart
      theirs = python_reads(@fixture.path, slide, probes)
      mismatches = probes.filter_map do |probe|
        ours = begin
          normalise(READ_PROBES.fetch(probe).call(chart))
        rescue StandardError => e
          "error:#{e.class}"
        end
        known = KNOWN_READ_DIVERGENCES[[slide, probe]]
        if known
          next if theirs[probe] == known[:python] && ours == known[:ruby]

          next "#{probe}: known divergence changed -- python #{theirs[probe].inspect}, ruby #{ours.inspect}"
        end
        "#{probe}: python #{theirs[probe].inspect}, ruby #{ours.inspect}" unless ours == theirs[probe]
      end
      expect(mismatches).to eq([])
    end
  end
end
