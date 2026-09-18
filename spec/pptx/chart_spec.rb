# frozen_string_literal: true

require "json"
require "open3"

RSpec.describe Pptx::Chart do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts[6]) }

  let(:chart_data) do
    Pptx::ChartData.new.tap do |data|
      data.categories = %w[East West Midwest]
      data.add_series("Q1", [1.2, 2.0, 3.5])
      data.add_series("Q2", [4.1, 5.0, 6.2])
    end
  end

  def add_chart(type = Pptx::Enum::XL_CHART_TYPE::COLUMN_CLUSTERED, data: chart_data)
    slide.shapes.add_chart(type, data, at: [Pptx.inches(1), Pptx.inches(1)],
                                       size: [Pptx.inches(8), Pptx.inches(5)])
  end

  describe "adding a chart" do
    subject(:frame) { add_chart }

    it "returns a graphic frame holding the chart" do
      aggregate_failures do
        expect(frame).to be_a(Pptx::GraphicFrame)
        expect(frame).to be_chart
        expect(frame).not_to be_table
        expect(frame.shape_type).to eq(Pptx::Enum::MSO_SHAPE_TYPE::CHART)
        expect(frame.name).to eq("Chart 1")
      end
    end

    it "raises when a frame is asked for a chart it does not hold" do
      table_frame = slide.shapes.add_table(1, 1, at: [0, 0], size: [100, 100])
      expect { table_frame.chart }.to raise_error(Pptx::Error, /does not contain a chart/)
    end

    it "creates a chart part and an embedded workbook part" do
      frame
      partnames = presentation.part.package.parts.map { |p| p.partname.to_s }
      aggregate_failures do
        expect(partnames).to include("/ppt/charts/chart1.xml")
        expect(partnames).to include("/ppt/embeddings/Microsoft_Excel_Sheet1.xlsx")
      end
    end

    it "reads back its own plot type, categories and values" do
      chart = frame.chart
      aggregate_failures do
        expect(chart.plot_type).to eq(:BAR)
        expect(chart.categories).to eq(%w[East West Midwest])
        expect(chart.series.map(&:name)).to eq(%w[Q1 Q2])
        expect(chart.series.first.values).to eq([1.2, 2.0, 3.5])
      end
    end

    it "records a gap rather than a zero for a missing value" do
      data = Pptx::ChartData.new
      data.categories = %w[A B C]
      data.add_series("S", [1, nil, 3])
      chart = add_chart(data: data).chart
      expect(chart.series.first.values).to eq([1.0, nil, 3.0])
    end

    # 27 of the 73 MS API chart types are supported; the 3-D ones and the
    # stock/surface families are not. Refusing beats writing XML PowerPoint
    # would reject.
    it "refuses a chart type it cannot write correctly" do
      expect { add_chart(Pptx::Enum::XL_CHART_TYPE::THREE_D_COLUMN) }
        .to raise_error(Pptx::Error, /not supported yet/)
    end

    it "creates a scatter chart from XY data" do
      data = Pptx::XyChartData.new
      data.add_series("Alpha", points: [[1, 10], [2, 20]])
      frame = slide.shapes.add_chart(:xy_scatter, data,
                                     at: [Pptx.inches(1), Pptx.inches(1)],
                                     size: [Pptx.inches(6), Pptx.inches(4)])
      aggregate_failures do
        expect(frame.chart.plot_type).to eq(:XY)
        expect(frame.chart.series.map(&:name)).to eq(["Alpha"])
      end
    end

    it "creates a bubble chart from bubble data" do
      data = Pptx::BubbleChartData.new
      data.add_series("B", points: [[1, 10, 5]])
      frame = slide.shapes.add_chart(:bubble, data,
                                     at: [Pptx.inches(1), Pptx.inches(1)],
                                     size: [Pptx.inches(6), Pptx.inches(4)])
      expect(frame.chart.plot_type).to eq(:BUBBLE)
    end
  end

  describe Pptx::ChartData do
    it "derives worksheet references from the data layout" do
      aggregate_failures do
        expect(chart_data.categories_ref).to eq("Sheet1!$A$2:$A$4")
        expect(chart_data.series[0].name_ref).to eq("Sheet1!$B$1")
        expect(chart_data.series[1].values_ref).to eq("Sheet1!$C$2:$C$4")
      end
    end

    it "names columns in Excel's bijective base-26" do
      aggregate_failures do
        expect(described_class.column_reference(1)).to eq("A")
        expect(described_class.column_reference(26)).to eq("Z")
        expect(described_class.column_reference(27)).to eq("AA")
        expect(described_class.column_reference(28)).to eq("AB")
        expect(described_class.column_reference(52)).to eq("AZ")
        expect(described_class.column_reference(53)).to eq("BA")
        expect(described_class.column_reference(702)).to eq("ZZ")
        expect(described_class.column_reference(703)).to eq("AAA")
      end
    end

    it "recognizes numeric categories" do
      numeric = described_class.new
      numeric.categories = [1, 2, 3]
      aggregate_failures do
        expect(numeric).to be_numeric_categories
        expect(chart_data).not_to be_numeric_categories
      end
    end
  end

  describe "the chart XML" do
    before { require_oracle! }

    # Ruby's nil is Python's None; everything else renders the same.
    def python_list(values)
      "(#{values.map { |v| v.nil? ? "None" : v.inspect }.join(", ")},)"
    end

    # Each family takes its own kind of data, so the fixture and the oracle
    # script are chosen to match.
    CATEGORIES = %w[East West Mid].freeze
    CATEGORY_SERIES = [["Q1", [1.2, 2.0, nil]], ["Q2", [4, 5, 6]]].freeze
    XY_SERIES = [["Alpha", [[1, 10], [2, 20]]], ["Beta", [[3, 30]]]].freeze
    BUBBLE_SERIES = [["B", [[1, 10, 5], [2, 20, 6]]]].freeze

    def ruby_data_for(family)
      case family
      when :xy
        Pptx::XyChartData.new.tap do |data|
          XY_SERIES.each { |name, points| data.add_series(name, points: points) }
        end
      when :bubble
        Pptx::BubbleChartData.new.tap do |data|
          BUBBLE_SERIES.each { |name, points| data.add_series(name, points: points) }
        end
      else
        Pptx::ChartData.new.tap do |data|
          data.categories = CATEGORIES
          CATEGORY_SERIES.each { |name, values| data.add_series(name, values) }
        end
      end
    end

    def python_data_for(family)
      case family
      when :xy
        lines = ["cd = XyChartData()"]
        XY_SERIES.each_with_index do |(name, points), i|
          lines << "s#{i} = cd.add_series(#{name.inspect})"
          points.each { |x, y| lines << "s#{i}.add_data_point(#{x}, #{y})" }
        end
        ["from pptx.chart.data import XyChartData", *lines].join("\n")
      when :bubble
        lines = ["cd = BubbleChartData()"]
        BUBBLE_SERIES.each_with_index do |(name, points), i|
          lines << "s#{i} = cd.add_series(#{name.inspect})"
          points.each { |x, y, size| lines << "s#{i}.add_data_point(#{x}, #{y}, #{size})" }
        end
        ["from pptx.chart.data import BubbleChartData", *lines].join("\n")
      else
        lines = ["from pptx.chart.data import CategoryChartData",
                 "cd = CategoryChartData()",
                 "cd.categories = #{CATEGORIES.inspect}"]
        CATEGORY_SERIES.each do |name, values|
          lines << "cd.add_series(#{name.inspect}, #{python_list(values)})"
        end
        lines.join("\n")
      end
    end

    def chart_xml_oracle(type_name, family)
      script = <<~PY
        import sys
        from pptx.enum.chart import XL_CHART_TYPE as XL
        #{python_data_for(family)}
        sys.stdout.write(cd.xml_bytes(getattr(XL, #{type_name.to_s.inspect})).decode())
      PY
      out, err, status = Open3.capture3("python3", "-c", script)
      raise "chart xml oracle failed: #{err}" unless status.success?

      out
    end

    # The cached values in this XML are what PowerPoint renders from, so it has
    # to match exactly. Every chart type this library can create is checked.
    Pptx::ChartXmlWriter::FAMILIES.each do |family, type_names|
      type_names.each do |type_name|
        it "matches python-pptx for #{type_name}" do
          ours = Pptx::ChartXmlWriter.write(Pptx::Enum::XL_CHART_TYPE.fetch(type_name),
                                            ruby_data_for(family))
          expect(ours).to eq(chart_xml_oracle(type_name, family))
        end
      end
    end
  end

  describe "the saved package" do
    before { require_oracle! }

    # The embedded workbook is deliberately excluded here: python-pptx writes
    # it with XlsxWriter and this gem writes it directly, so the bytes differ
    # by construction. Everything else -- including the chart XML PowerPoint
    # actually renders -- must match exactly, and the workbook's contents are
    # checked separately below.
    it "matches python-pptx except for the embedded workbook" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          from pptx.chart.data import CategoryChartData
          from pptx.enum.chart import XL_CHART_TYPE
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          cd = CategoryChartData()
          cd.categories = ["East", "West", "Midwest"]
          cd.add_series("Q1", (1.2, 2.0, 3.5))
          cd.add_series("Q2", (4.1, 5.0, 6.2))
          slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1),
                                 Inches(8), Inches(5), cd)
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          data = Pptx::ChartData.new
          data.categories = %w[East West Midwest]
          data.add_series("Q1", [1.2, 2.0, 3.5])
          data.add_series("Q2", [4.1, 5.0, 6.2])
          slide.shapes.add_chart(Pptx::Enum::XL_CHART_TYPE::COLUMN_CLUSTERED, data,
                                 at: [Pptx.inches(1), Pptx.inches(1)], size: [Pptx.inches(8), Pptx.inches(5)])
          prs.save(path)
        },
        ignore: ["ppt/embeddings/Microsoft_Excel_Sheet1.xlsx"]
      )
    end
  end

  describe "what python-pptx sees in our output" do
    def chart_oracle_path = File.expand_path("../../tools/chart_oracle.py", __dir__)

    before do
      require_oracle!
      require_oracle!("openpyxl", available: Pptx::Spec::Differential.openpyxl_available?)
    end

    # End-to-end: our file is opened by python-pptx, and its embedded workbook
    # by a real spreadsheet reader. This is what makes the workbook exclusion
    # above safe -- the grid is checked, just not the bytes.
    it "round-trips the chart and its workbook through python-pptx" do
      add_chart
      result = Tempfile.create(["chart", ".pptx"]) do |file|
        file.close
        presentation.save(file.path)
        out, err, status = Open3.capture3("python3", chart_oracle_path, file.path)
        raise "chart oracle failed: #{err}" unless status.success?

        JSON.parse(out)
      end

      chart = result["charts"].first
      aggregate_failures do
        expect(chart["chart_type"]).to eq("COLUMN_CLUSTERED (51)")
        expect(chart["categories"]).to eq(%w[East West Midwest])
        expect(chart["series"]).to eq(
          [{ "name" => "Q1", "values" => [1.2, 2.0, 3.5] },
           { "name" => "Q2", "values" => [4.1, 5.0, 6.2] }]
        )
        expect(chart["sheet_name"]).to eq("Sheet1")
        expect(chart["grid"]).to eq(
          [[nil, "Q1", "Q2"],
           ["East", 1.2, 4.1],
           ["West", 2, 5],
           ["Midwest", 3.5, 6.2]]
        )
      end
    end
  end
end
