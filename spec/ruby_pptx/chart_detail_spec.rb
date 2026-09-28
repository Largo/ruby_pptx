# frozen_string_literal: true

# Writing the chart detail python-pptx supports -- axes, legend, labels,
# markers, points, fonts -- must produce the same package it does.
RSpec.describe "chart detail agreement with python-pptx" do
  before { require_oracle! }

  CHART_PY_HEADER = <<~PY
    from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData
    from pptx.enum.chart import (XL_CHART_TYPE as XL, XL_LEGEND_POSITION, XL_TICK_MARK,
        XL_TICK_LABEL_POSITION, XL_AXIS_CROSSES, XL_DATA_LABEL_POSITION, XL_MARKER_STYLE)
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt
    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cd = CategoryChartData()
    cd.categories = ["East", "West", "Mid"]
    cd.add_series("Q1", (1.5, -2, 3))
    cd.add_series("Q2", (4, 5, 6.25))
  PY

  def ruby_chart(type, path)
    prs = Pptx::Presentation.new_default
    slide = prs.slides.add(prs.slide_layouts[6])
    data = Pptx::ChartData.new
    data.categories = %w[East West Mid]
    data.add_series("Q1", [1.5, -2, 3])
    data.add_series("Q2", [4, 5, 6.25])
    chart = slide.shapes.add_chart(type, data, at: [0, 0],
                                               size: [Pptx.inches(6), Pptx.inches(4)]).chart
    yield chart
    prs.save(path)
  end

  def python_chart(type, body)
    CHART_PY_HEADER + <<~PY
      chart = slide.shapes.add_chart(XL.#{type}, 0, 0, Inches(6), Inches(4), cd).chart
      #{body}
      prs.save(out)
    PY
  end

  # The workbook is written by XlsxWriter on one side and directly on the
  # other; its contents are checked elsewhere.
  CHART_WORKBOOK = ["ppt/embeddings/Microsoft_Excel_Sheet1.xlsx"].freeze

  it "writes axis settings identically" do
    expect_same_package(
      python: python_chart("BAR_STACKED", <<~PY),
        ca, va = chart.category_axis, chart.value_axis
        ca.major_tick_mark = XL_TICK_MARK.OUTSIDE
        ca.minor_tick_mark = XL_TICK_MARK.INSIDE
        ca.reverse_order = True
        ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        ca.has_major_gridlines = True
        ca.major_gridlines.format.line.width = Pt(2)
        ca.tick_labels.offset = 250
        ca.tick_labels.font.size = Pt(9)
        ca.tick_labels.font.bold = True
        ca.format.line.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        ca.has_title = True
        va.maximum_scale = 12.5
        va.minimum_scale = -3
        va.major_unit = 2.5
        va.minor_unit = 0.5
        va.crosses = XL_AXIS_CROSSES.MAXIMUM
        va.reverse_order = True
        va.reverse_order = False
        va.tick_labels.number_format = "0.0%"
        ca.tick_labels.number_format_is_linked = True
        va.axis_title.text_frame.text = "Units"
        va.major_tick_mark = XL_TICK_MARK.CROSS
        va.visible = False
      PY
      ruby: lambda { |path|
        ruby_chart(:bar_stacked, path) do |chart|
          ca = chart.category_axis
          va = chart.value_axis
          ca.major_tick_mark = :outside
          ca.minor_tick_mark = :inside
          ca.reverse_order = true
          ca.tick_label_position = :low
          ca.major_gridlines = true
          ca.major_gridlines.format.line.width = Pptx.pt(2)
          ca.tick_labels.offset = 250
          ca.tick_labels.font.size = Pptx.pt(9)
          ca.tick_labels.font.bold = true
          ca.format.line.color.rgb = Pptx::RGBColor["1F497D"]
          ca.title = true
          va.maximum_scale = 12.5
          va.minimum_scale = -3
          va.major_unit = 2.5
          va.minor_unit = 0.5
          va.crosses = :maximum
          va.reverse_order = true
          va.reverse_order = false
          va.tick_labels.number_format = "0.0%"
          ca.tick_labels.number_format_linked = true
          va.title = "Units"
          va.major_tick_mark = :cross
          va.visible = false
        end
      },
      ignore: CHART_WORKBOOK
    )
  end

  it "writes a custom crossing point identically" do
    expect_same_package(
      python: python_chart("COLUMN_CLUSTERED", <<~PY),
        chart.value_axis.crosses_at = 1.5
        chart.value_axis.crosses = XL_AXIS_CROSSES.CUSTOM
        chart.value_axis.crosses = XL_AXIS_CROSSES.MINIMUM
        chart.value_axis.crosses = XL_AXIS_CROSSES.CUSTOM
        chart.category_axis.tick_labels.offset = 100
      PY
      ruby: lambda { |path|
        ruby_chart(:column_clustered, path) do |chart|
          chart.value_axis.crosses_at = 1.5
          chart.value_axis.crosses = :custom
          # From a named crossing to a custom one, which has to drop c:crosses.
          chart.value_axis.crosses = :minimum
          chart.value_axis.crosses = :custom
          chart.category_axis.tick_labels.offset = 100
        end
      },
      ignore: CHART_WORKBOOK
    )
  end

  it "writes chart-level settings, legend and title identically" do
    expect_same_package(
      python: python_chart("COLUMN_CLUSTERED", <<~PY),
        chart.chart_style = 17
        chart.font.size = Pt(11)
        chart.font.bold = True
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.horz_offset = 0.2
        chart.legend.font.size = Pt(9)
        chart.chart_title.text_frame.text = "Revenue"
        chart.chart_title.format.fill.solid()
      PY
      ruby: lambda { |path|
        ruby_chart(:column_clustered, path) do |chart|
          chart.chart_style = 17
          chart.font.size = Pptx.pt(11)
          chart.font.bold = true
          chart.legend = true
          chart.legend.position = :bottom
          chart.legend.include_in_layout = false
          chart.legend.horz_offset = 0.2
          chart.legend.font.size = Pptx.pt(9)
          chart.title = "Revenue"
          chart.title.format.fill.solid
        end
      },
      ignore: CHART_WORKBOOK
    )
  end

  # Removing a title must also record that it was deleted, or PowerPoint
  # regrows a generated one on a single-series chart.
  it "removes a title identically" do
    expect_same_package(
      python: python_chart("COLUMN_CLUSTERED", <<~PY),
        chart.chart_style = 17
        chart.chart_style = None
        chart.has_title = True
        chart.has_title = False
        chart.has_legend = True
        chart.legend.horz_offset = 0.3
        chart.legend.horz_offset = 0.0
      PY
      ruby: lambda { |path|
        ruby_chart(:column_clustered, path) do |chart|
          chart.chart_style = 17
          chart.chart_style = nil
          chart.title = true
          chart.title = nil
          chart.legend = true
          chart.legend.horz_offset = 0.3
          chart.legend.horz_offset = 0.0
        end
      },
      ignore: CHART_WORKBOOK
    )
  end

  it "writes data labels identically" do
    expect_same_package(
      python: python_chart("COLUMN_CLUSTERED", <<~PY),
        p = chart.plots[0]
        p.has_data_labels = True
        p.gap_width = 80
        p.overlap = 0
        p.vary_by_categories = False
        dl = p.data_labels
        dl.number_format = "#,##0"
        dl.position = XL_DATA_LABEL_POSITION.INSIDE_END
        dl.font.size = Pt(8)
        dl.show_category_name = True
        s = chart.series[1]
        s.invert_if_negative = False
        s.data_labels.show_value = True
        s.data_labels.number_format_is_linked = False
      PY
      ruby: lambda { |path|
        ruby_chart(:column_clustered, path) do |chart|
          plot = chart.plots[0]
          plot.data_labels = true
          plot.gap_width = 80
          plot.overlap = 0
          plot.vary_by_categories = false
          labels = plot.data_labels
          labels.number_format = "#,##0"
          labels.position = :inside_end
          labels.font.size = Pptx.pt(8)
          labels.show_category_name = true
          series = chart.series[1]
          series.invert_if_negative = false
          series.data_labels.show_value = true
          series.data_labels.number_format_linked = false
        end
      },
      ignore: CHART_WORKBOOK
    )
  end

  # Points are the finest grain: a label, a fill or a marker on one point.
  # Labels are inserted in index order, so they are added out of order here.
  it "writes per-point formatting identically" do
    expect_same_package(
      python: python_chart("LINE_MARKERS", <<~PY),
        s = chart.series[0]
        s.smooth = False
        s.marker.style = XL_MARKER_STYLE.DIAMOND
        s.marker.size = 5
        s.marker.size = 11
        s.marker.format.fill.solid()
        s.marker.format.fill.fore_color.rgb = RGBColor(0xC0, 0x50, 0x4D)
        chart.series[1].smooth = True
        chart.series[1].marker.size = 5
        chart.series[1].marker.size = None
        p2 = s.points[2]
        p2.data_label.position = XL_DATA_LABEL_POSITION.ABOVE
        p0 = s.points[0]
        p0.data_label.text_frame.text = "start"
        p0.data_label.font.italic = True
        p1 = s.points[1]
        p1.marker.style = XL_MARKER_STYLE.SQUARE
        p1.marker.size = 7
        p1.format.line.width = Pt(3)
        p1.data_label.has_text_frame = True
        p1.data_label.position = XL_DATA_LABEL_POSITION.BELOW
      PY
      ruby: lambda { |path|
        ruby_chart(:line_markers, path) do |chart|
          series = chart.series[0]
          series.smooth = false
          series.marker.style = :diamond
          series.marker.size = 5
          series.marker.size = 11
          series.marker.format.fill.solid
          series.marker.format.fill.fore_color.rgb = Pptx::RGBColor["C0504D"]
          chart.series[1].smooth = true
          chart.series[1].marker.size = 5
          chart.series[1].marker.size = nil
          series.points[2].data_label.position = :above
          series.points[0].data_label.text_frame.text = "start"
          series.points[0].data_label.font.italic = true
          point = series.points[1]
          point.marker.style = :square
          point.marker.size = 7
          point.format.line.width = Pptx.pt(3)
          point.data_label.text_frame
          point.data_label.position = :below
        end
      },
      ignore: CHART_WORKBOOK
    )
  end

  it "writes a bubble scale identically" do
    expect_same_package(
      python: <<~PY,
        from pptx.chart.data import BubbleChartData
        from pptx.enum.chart import XL_CHART_TYPE as XL
        from pptx.util import Inches
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bd = BubbleChartData()
        s = bd.add_series("B")
        for x, y, z in ((1, 10, 5), (2, 20, 7)):
            s.add_data_point(x, y, z)
        chart = slide.shapes.add_chart(XL.BUBBLE, 0, 0, Inches(6), Inches(4), bd).chart
        chart.plots[0].bubble_scale = 150
        chart.series[0].points[1].format.fill.background()
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        data = Pptx::BubbleChartData.new
        data.add_series("B", points: [[1, 10, 5], [2, 20, 7]])
        chart = slide.shapes.add_chart(:bubble, data, at: [0, 0],
                                                      size: [Pptx.inches(6), Pptx.inches(4)]).chart
        chart.plots[0].bubble_scale = 150
        chart.series[0].points[1].format.fill.background
        prs.save(path)
      },
      ignore: CHART_WORKBOOK
    )
  end
end
