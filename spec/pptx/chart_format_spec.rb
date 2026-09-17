# frozen_string_literal: true

RSpec.describe "chart formatting" do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }

  def chart_of(type = :column_clustered)
    data = Pptx::ChartData.new
    data.categories = %w[East West Mid]
    data.add_series("Q1", [1, 2, 3])
    data.add_series("Q2", [4, 5, 6])
    slide.shapes.add_chart(type, data, at: [Pptx.inches(1), Pptx.inches(1)],
                                       size: [Pptx.inches(8), Pptx.inches(5)]).chart
  end

  subject(:chart) { chart_of }

  describe "the legend" do
    it "reflects whether the chart type writes one by default" do
      aggregate_failures do
        expect(chart_of(:column_clustered).legend?).to be(false)
        expect(chart_of(:line).legend?).to be(true)
      end
    end

    it "adds and removes one" do
      aggregate_failures do
        chart.legend = true
        expect(chart.legend?).to be(true)
        expect(chart.legend).to be_a(Pptx::ChartLegend)

        chart.legend = false
        expect(chart.legend?).to be(false)
        expect(chart.legend).to be_nil
      end
    end

    it "round-trips position and overlay" do
      chart.legend = true
      chart.legend.position = :BOTTOM
      chart.legend.include_in_layout = false
      aggregate_failures do
        expect(chart.legend.position).to eq(Pptx::Enum::XL_LEGEND_POSITION::BOTTOM)
        expect(chart.legend.include_in_layout?).to be(false)
        expect(chart.element.xml).to include('<c:legendPos val="b"/>')
      end
    end

    it "accepts a position by symbol, ignoring case" do
      chart.legend = true
      chart.legend.position = :right
      expect(chart.legend.position).to eq(Pptx::Enum::XL_LEGEND_POSITION::RIGHT)
    end
  end

  describe "the chart title" do
    it "is absent until set" do
      aggregate_failures do
        expect(chart.title?).to be(false)
        expect(chart.title).to be_nil
      end
    end

    it "round-trips text" do
      chart.title = "Quarterly Revenue"
      aggregate_failures do
        expect(chart.title?).to be(true)
        expect(chart.title.text).to eq("Quarterly Revenue")
      end
    end

    it "replaces rather than appending when set twice" do
      chart.title = "First"
      chart.title = "Second"
      aggregate_failures do
        expect(chart.title.text).to eq("Second")
        expect(chart.element.xml.scan(/<c:title>/).size).to eq(1)
      end
    end

    it "removes with nil" do
      chart.title = "Gone"
      chart.title = nil
      aggregate_failures do
        expect(chart.title?).to be(false)
        expect(chart.element.xml).not_to include("<c:title>")
      end
    end

    # A title is a text frame like any other, so it formats the same way.
    it "exposes a text frame for formatting" do
      chart.title = "Styled"
      chart.title.text_frame.paragraphs.first.runs.first.font.bold = true
      expect(chart.element.xml).to include('b="1"')
    end
  end

  describe "axes" do
    it "gives a bar chart a category axis and a value axis" do
      aggregate_failures do
        expect(chart.category_axis).to be_a(Pptx::ChartAxis)
        expect(chart.value_axis).to be_a(Pptx::ChartAxis)
      end
    end

    # A pie has neither axis; asking should give nil rather than raising.
    it "reports no axes for a pie" do
      pie = chart_of(:pie)
      aggregate_failures do
        expect(pie.category_axis).to be_nil
        expect(pie.value_axis).to be_nil
      end
    end

    it "gives a scatter chart two value axes" do
      data = Pptx::XyChartData.new
      data.add_series("A", points: [[1, 2]])
      xy = slide.shapes.add_chart(:xy_scatter, data, at: [0, 0],
                                                     size: [Pptx.inches(4), Pptx.inches(3)]).chart
      expect(xy.value_axes.size).to eq(2)
    end

    it "round-trips the scale" do
      axis = chart.value_axis
      axis.maximum_scale = 10.0
      axis.minimum_scale = 2.0
      axis.major_unit = 2.0
      aggregate_failures do
        expect(axis.maximum_scale).to eq(10.0)
        expect(axis.minimum_scale).to eq(2.0)
        expect(axis.major_unit).to eq(2.0)
      end
    end

    # An absent element means "let PowerPoint decide", which is what nil says.
    it "reads an unset scale as nil, meaning automatic" do
      axis = chart.value_axis
      aggregate_failures do
        expect(axis.maximum_scale).to be_nil
        axis.maximum_scale = 10.0
        axis.maximum_scale = nil
        expect(axis.maximum_scale).to be_nil
        expect(axis.element.xml).not_to include("<c:max")
      end
    end

    it "toggles gridlines" do
      axis = chart.category_axis
      aggregate_failures do
        expect(axis.major_gridlines?).to be(false)
        axis.major_gridlines = true
        expect(axis.major_gridlines?).to be(true)
        axis.major_gridlines = false
        expect(axis.major_gridlines?).to be(false)
      end
    end

    it "hides and shows the axis" do
      axis = chart.value_axis
      aggregate_failures do
        expect(axis.visible?).to be(true)
        axis.visible = false
        expect(axis.visible?).to be(false)
        expect(axis.element.xml).to include('<c:delete val="1"/>')
      end
    end

    # Setting an explicit format means the labels no longer follow the source
    # data, which is what sourceLinked records.
    it "sets a number format and unlinks it from the source" do
      axis = chart.value_axis
      axis.number_format = "0.0%"
      aggregate_failures do
        expect(axis.number_format).to eq("0.0%")
        expect(axis.element.numFmt.sourceLinked).to be(false)
      end
    end

    it "round-trips an axis title" do
      axis = chart.value_axis
      axis.title = "Millions"
      aggregate_failures do
        expect(axis.has_title?).to be(true)
        expect(axis.title.text).to eq("Millions")
        axis.title = nil
        expect(axis.has_title?).to be(false)
      end
    end
  end

  describe "plots" do
    it "reports one plot for an ordinary chart" do
      aggregate_failures do
        expect(chart.plots.size).to eq(1)
        expect(chart.plots.first).to be_a(Pptx::ChartPlot)
      end
    end

    it "round-trips gap width and overlap" do
      plot = chart.plots.first
      aggregate_failures do
        expect(plot.gap_width).to eq(150)
        expect(plot.overlap).to eq(0)
        plot.gap_width = 75
        plot.overlap = -10
        expect(plot.gap_width).to eq(75)
        expect(plot.overlap).to eq(-10)
      end
    end

    it "reports vary-by-categories, which a pie sets and a bar does not" do
      aggregate_failures do
        expect(chart.plots.first.vary_by_categories?).to be(false)
        expect(chart_of(:pie).plots.first.vary_by_categories?).to be(true)
      end
    end
  end

  describe "data labels" do
    subject(:labels) { chart.plots.first.data_labels }

    it "switch on and off as a whole" do
      plot = chart.plots.first
      aggregate_failures do
        expect(plot.data_labels?).to be(false)
        plot.data_labels.show_value = true
        expect(plot.data_labels?).to be(true)
        plot.data_labels = false
        expect(plot.data_labels?).to be(false)
      end
    end

    # PowerPoint writes every label kind explicitly when labels are switched
    # on; an empty c:dLbls would not satisfy the schema.
    it "are created with the full default set of flags" do
      labels.show_value = true
      xml = chart.element.xml
      aggregate_failures do
        %w[showLegendKey showVal showCatName showSerName showPercent
           showBubbleSize showLeaderLines].each do |flag|
          expect(xml).to include("<c:#{flag} ")
        end
      end
    end

    it "are all off until switched on" do
      aggregate_failures do
        expect(labels.show_value?).to be(false)
        expect(labels.show_category_name?).to be(false)
        expect(labels.show_percentage?).to be(false)
      end
    end

    it "round-trips each flag" do
      labels.show_value = true
      labels.show_series_name = true
      aggregate_failures do
        expect(labels.show_value?).to be(true)
        expect(labels.show_series_name?).to be(true)
        expect(labels.show_category_name?).to be(false)
      end
    end

    it "round-trips number format and position" do
      labels.number_format = "0.0"
      labels.position = :OUTSIDE_END
      aggregate_failures do
        expect(labels.number_format).to eq("0.0")
        expect(labels.position).to eq(Pptx::Enum::XL_LABEL_POSITION::OUTSIDE_END)
        labels.position = nil
        expect(labels.position).to be_nil
      end
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "formats a chart identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          from pptx.chart.data import CategoryChartData
          from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          cd = CategoryChartData()
          cd.categories = ["East", "West", "Mid"]
          cd.add_series("Q1", (1, 2, 3))
          cd.add_series("Q2", (4, 5, 6))
          chart = slide.shapes.add_chart(
              XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1),
              Inches(8), Inches(5), cd).chart

          chart.has_title = True
          chart.chart_title.text_frame.text = "Quarterly Revenue"
          chart.has_legend = True
          chart.legend.position = XL_LEGEND_POSITION.BOTTOM
          chart.legend.include_in_layout = False

          va = chart.value_axis
          va.maximum_scale = 10.0
          va.minimum_scale = 0.0
          va.major_unit = 2.0
          va.tick_labels.number_format = "0.0"
          va.tick_labels.number_format_is_linked = False
          va.has_title = True
          va.axis_title.text_frame.text = "Millions"

          chart.category_axis.has_major_gridlines = True

          plot = chart.plots[0]
          plot.gap_width = 75
          plot.overlap = -10
          plot.has_data_labels = True
          plot.data_labels.show_value = True
          plot.data_labels.number_format = "0.0"
          plot.data_labels.number_format_is_linked = False
          plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          data = Pptx::ChartData.new
          data.categories = ["East", "West", "Mid"]
          data.add_series("Q1", [1, 2, 3])
          data.add_series("Q2", [4, 5, 6])
          chart = slide.shapes.add_chart(:column_clustered, data,
                                         at: [Pptx.inches(1), Pptx.inches(1)],
                                         size: [Pptx.inches(8), Pptx.inches(5)]).chart

          chart.title = "Quarterly Revenue"
          chart.legend = true
          chart.legend.position = :BOTTOM
          chart.legend.include_in_layout = false

          axis = chart.value_axis
          axis.maximum_scale = 10.0
          axis.minimum_scale = 0.0
          axis.major_unit = 2.0
          axis.number_format = "0.0"
          axis.title = "Millions"

          chart.category_axis.major_gridlines = true

          plot = chart.plots.first
          plot.gap_width = 75
          plot.overlap = -10
          plot.data_labels.show_value = true
          plot.data_labels.number_format = "0.0"
          plot.data_labels.position = :OUTSIDE_END
          prs.save(path)
        },
        ignore: ["ppt/embeddings/Microsoft_Excel_Sheet1.xlsx"]
      )
    end
  end
end
