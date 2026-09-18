# frozen_string_literal: true

RSpec.describe "chart series formatting and data replacement" do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }

  SIX_BY_FOUR = [Pptx.inches(6), Pptx.inches(4)].freeze

  def data_of(categories, series)
    Pptx::ChartData.new.tap do |d|
      d.categories = categories
      series.each { |name, values| d.add_series(name, values) }
    end
  end

  let(:chart) do
    slide.shapes.add_chart(:column_clustered,
                           data_of(%w[East West], [["Q1", [1, 2]], ["Q2", [3, 4]]]),
                           at: [0, 0], size: [Pptx.inches(6), Pptx.inches(4)]).chart
  end

  describe "series formatting" do
    it "gives each series its own fill and line" do
      first = chart.series.first
      first.format.fill.solid
      first.format.fill.fore_color.rgb = Pptx::RGBColor["C0504D"]
      first.format.line.color.rgb = Pptx::RGBColor["1F497D"]

      aggregate_failures do
        expect(first.format.fill.type).to eq(Pptx::Enum::MSO_FILL_TYPE::SOLID)
        expect(first.element.xml).to include("C0504D", "1F497D")
        # The other series is untouched.
        expect(chart.series.last.element.xml).not_to include("C0504D")
      end
    end

    it "puts the formatting in the series' own shape properties" do
      chart.series.first.format.fill.background
      expect(chart.series.first.element.xpath("./c:spPr/a:noFill")).not_to be_empty
    end
  end

  describe "#replace_data" do
    it "replaces categories and values" do
      chart.replace_data(data_of(%w[North South Central],
                                 [["Q1", [10, 20, 30]], ["Q2", [40, 50, 60]]]))
      aggregate_failures do
        expect(chart.categories).to eq(%w[North South Central])
        expect(chart.series.map(&:values)).to eq([[10.0, 20.0, 30.0], [40.0, 50.0, 60.0]])
      end
    end

    it "renames series" do
      chart.replace_data(data_of(%w[A], [["Alpha", [1]], ["Beta", [2]]]))
      expect(chart.series.map(&:name)).to eq(%w[Alpha Beta])
    end

    # The point of replacing data rather than rebuilding is to keep the work
    # already done on appearance.
    it "leaves series formatting alone" do
      chart.series.first.format.fill.solid
      chart.series.first.format.fill.fore_color.rgb = Pptx::RGBColor["C0504D"]
      chart.replace_data(data_of(%w[A B], [["Q1", [9, 8]], ["Q2", [7, 6]]]))
      expect(chart.series.first.element.xml).to include("C0504D")
    end

    it "adds series when the new data has more" do
      chart.replace_data(data_of(%w[A], [["One", [1]], ["Two", [2]], ["Three", [3]]]))
      aggregate_failures do
        expect(chart.series.map(&:name)).to eq(%w[One Two Three])
        expect(chart.element.xpath("//c:ser/c:idx/@val").map { |a| a.value.to_i }).to eq([0, 1, 2])
      end
    end

    # A new series copies the last one, so it arrives looking like its
    # neighbour rather than unstyled.
    it "clones formatting onto series it adds" do
      chart.series.last.format.line.color.rgb = Pptx::RGBColor["1F497D"]
      chart.replace_data(data_of(%w[A], [["One", [1]], ["Two", [2]], ["Three", [3]]]))
      expect(chart.series.last.element.xml).to include("1F497D")
    end

    it "removes series when the new data has fewer" do
      chart.replace_data(data_of(%w[A], [["Only", [1]]]))
      expect(chart.series.map(&:name)).to eq(["Only"])
    end

    # An empty plot element has nothing to draw and is not valid, so it goes
    # when its last series does.
    it "removes a plot left with no series" do
      combo = slide.shapes.add_combo_chart(
        data_of(%w[A B], [["Bars", [1, 2]], ["Line", [3, 4]]]),
        at: [0, 0], size: [Pptx.inches(6), Pptx.inches(4)]
      ) do |c|
        c.plot :column_clustered, series: "Bars"
        c.plot :line, series: "Line"
      end.chart

      expect(combo.plots.size).to eq(2)
      combo.replace_data(data_of(%w[A], [["Bars", [1]]]))
      aggregate_failures do
        expect(combo.plots.size).to eq(1)
        expect(combo.plots.first.element.nsptag).to eq("c:barChart")
      end
    end

    it "rewrites the embedded workbook too" do
      chart.replace_data(data_of(%w[North South], [["Q1", [10, 20]]]))
      expect(chart.part.workbook.xlsx_part.blob).not_to be_empty
    end

    it "refuses to add series to a chart that has none" do
      empty = chart
      empty.element.xpath("//c:ser").each { |ser| ser.parent.remove(ser) }
      expect { empty.replace_data(data_of(%w[A], [["X", [1]]])) }
        .to raise_error(Pptx::Error, /chart that has none/)
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "formats series identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches, Pt
          from pptx.chart.data import CategoryChartData
          from pptx.enum.chart import XL_CHART_TYPE
          from pptx.dml.color import RGBColor
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          cd = CategoryChartData()
          cd.categories = ["East", "West"]
          cd.add_series("Q1", (1, 2))
          cd.add_series("Q2", (3, 4))
          chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 0,
                                         Inches(6), Inches(4), cd).chart
          s0 = chart.series[0]
          s0.format.fill.solid()
          s0.format.fill.fore_color.rgb = RGBColor(0xC0, 0x50, 0x4D)
          s1 = chart.series[1]
          s1.format.line.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
          s1.format.line.width = Pt(2)
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          data = Pptx::ChartData.new
          data.categories = ["East", "West"]
          data.add_series("Q1", [1, 2])
          data.add_series("Q2", [3, 4])
          chart = slide.shapes.add_chart(:column_clustered, data, at: [0, 0],
                                                                  size: SIX_BY_FOUR).chart
          first = chart.series[0]
          first.format.fill.solid
          first.format.fill.fore_color.rgb = Pptx::RGBColor.new(0xC0, 0x50, 0x4D)
          second = chart.series[1]
          second.format.line.color.rgb = Pptx::RGBColor.new(0x1F, 0x49, 0x7D)
          second.format.line.width = Pptx.pt(2)
          prs.save(path)
        },
        ignore: ["ppt/embeddings/Microsoft_Excel_Sheet1.xlsx"]
      )
    end

    it "replaces data identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          from pptx.chart.data import CategoryChartData
          from pptx.enum.chart import XL_CHART_TYPE
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          cd = CategoryChartData()
          cd.categories = ["East", "West"]
          cd.add_series("Q1", (1, 2))
          cd.add_series("Q2", (3, 4))
          chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 0,
                                         Inches(6), Inches(4), cd).chart

          fresh = CategoryChartData()
          fresh.categories = ["North", "South", "Central"]
          fresh.add_series("A", (10, 20, 30))
          fresh.add_series("B", (40, 50, 60))
          fresh.add_series("C", (70, 80, 90))
          chart.replace_data(fresh)
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          data = Pptx::ChartData.new
          data.categories = ["East", "West"]
          data.add_series("Q1", [1, 2])
          data.add_series("Q2", [3, 4])
          chart = slide.shapes.add_chart(:column_clustered, data, at: [0, 0],
                                                                  size: SIX_BY_FOUR).chart

          fresh = Pptx::ChartData.new
          fresh.categories = %w[North South Central]
          fresh.add_series("A", [10, 20, 30])
          fresh.add_series("B", [40, 50, 60])
          fresh.add_series("C", [70, 80, 90])
          chart.replace_data(fresh)
          prs.save(path)
        },
        ignore: ["ppt/embeddings/Microsoft_Excel_Sheet1.xlsx"]
      )
    end

    it "replaces data down to fewer series identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          from pptx.chart.data import CategoryChartData
          from pptx.enum.chart import XL_CHART_TYPE
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          cd = CategoryChartData()
          cd.categories = ["East", "West"]
          cd.add_series("Q1", (1, 2))
          cd.add_series("Q2", (3, 4))
          chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, 0, 0,
                                         Inches(6), Inches(4), cd).chart
          fewer = CategoryChartData()
          fewer.categories = ["Only"]
          fewer.add_series("Solo", (7,))
          chart.replace_data(fewer)
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          data = Pptx::ChartData.new
          data.categories = ["East", "West"]
          data.add_series("Q1", [1, 2])
          data.add_series("Q2", [3, 4])
          chart = slide.shapes.add_chart(:column_clustered, data, at: [0, 0],
                                                                  size: SIX_BY_FOUR).chart
          fewer = Pptx::ChartData.new
          fewer.categories = ["Only"]
          fewer.add_series("Solo", [7])
          chart.replace_data(fewer)
          prs.save(path)
        },
        ignore: ["ppt/embeddings/Microsoft_Excel_Sheet1.xlsx"]
      )
    end
  end
end
