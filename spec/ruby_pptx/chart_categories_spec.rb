# frozen_string_literal: true

require "date"
require "json"

RSpec.describe "chart categories: grouped, dated and numeric" do
  describe Pptx::ChartDataCategories do
    it "takes a flat list" do
      categories = described_class.from(%w[East West])
      aggregate_failures do
        expect(categories.depth).to eq(1)
        expect(categories.leaf_count).to eq(2)
        expect(categories.leaf_labels).to eq(%w[East West])
      end
    end

    it "takes groups as a Hash, nested as deep as needed" do
      categories = described_class.from("East" => { "North" => %w[A B], "South" => %w[C] },
                                        "West" => { "Coast" => %w[D] })
      aggregate_failures do
        expect(categories.depth).to eq(3)
        expect(categories.leaf_count).to eq(4)
        expect(categories.leaf_labels).to eq(%w[A B C D])
      end
    end

    it "builds the same tree one category at a time" do
      categories = described_class.new
      east = categories.add_category("East")
      east.add_sub_category("North")
      east.add_sub_category("South")
      categories.add_category("West").add_sub_category("Coast")
      expect(categories.levels).to eq(described_class.from("East" => %w[North South],
                                                           "West" => %w[Coast]).levels)
    end

    # A parent's index is where its run of leaves begins.
    it "lists each level leaf first, with parents at the start of their runs" do
      categories = described_class.from("East" => %w[North South], "West" => %w[Coast])
      expect(categories.levels).to eq([[[0, "North"], [1, "South"], [2, "Coast"]],
                                       [[0, "East"], [2, "West"]]])
    end

    it "refuses groups of uneven depth" do
      categories = described_class.from("East" => %w[North], "West" => nil)
      expect { categories.depth }.to raise_error(Pptx::Error, /not uniform/)
    end

    # As python-pptx decides it: from the first category.
    it "recognises dates and numbers by the first category" do
      aggregate_failures do
        expect(described_class.from([Date.new(2024, 1, 1)])).to be_dates
        expect(described_class.from([Date.new(2024, 1, 1)])).to be_numeric
        expect(described_class.from([1, 2])).to be_numeric
        expect(described_class.from([1, 2])).not_to be_dates
        expect(described_class.from(%w[a b])).not_to be_numeric
        expect(described_class.from("G" => [1, 2])).not_to be_numeric
      end
    end

    it "formats dates as dates unless told otherwise" do
      categories = described_class.from([Date.new(2024, 1, 1)])
      aggregate_failures do
        expect(categories.number_format).to eq("yyyy\\-mm\\-dd")
        categories.number_format = "mmm yy"
        expect(categories.number_format).to eq("mmm yy")
        expect(described_class.from(%w[a]).number_format).to eq("General")
      end
    end

    it "gives a nil label as an empty string" do
      expect(described_class.from(["a", nil]).leaf_labels).to eq(["a", ""])
    end
  end

  describe Pptx::ChartDataCategory do
    # 1900 is treated as a leap year by Excel, so every serial after 28
    # February 1900 is one more than the day count; 45306 is 2024-01-15.
    it "writes a date as its Excel serial number" do
      category = Pptx::ChartDataCategories.from([Date.new(2024, 1, 15)]).first
      aggregate_failures do
        expect(category.numeric_str_val).to eq("45306.0")
        expect(category.numeric_str_val(date_1904: true)).to eq("43844.0")
      end
    end

    it "counts from the epoch without the leap-year correction before March 1900" do
      category = Pptx::ChartDataCategories.from([Date.new(1900, 2, 28)]).first
      expect(category.numeric_str_val).to eq("59.0")
    end
  end

  describe "the worksheet references" do
    it "spans one column per category level, series after them" do
      data = Pptx::ChartData.new
      data.categories = { "East" => %w[North South], "West" => %w[Coast] }
      data.add_series("S", [1, 2, 3])
      aggregate_failures do
        expect(data.categories_ref).to eq("Sheet1!$A$2:$B$4")
        expect(data.series[0].name_ref).to eq("Sheet1!$C$1")
        expect(data.series[0].values_ref).to eq("Sheet1!$C$2:$C$4")
      end
    end

    # python-pptx measures a series' range by the series, not the categories.
    it "ends a series' range at its own length" do
      data = Pptx::ChartData.new
      data.categories = %w[a b c]
      data.add_series("S", [1, 2])
      expect(data.series[0].values_ref).to eq("Sheet1!$B$2:$B$3")
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    def python_xml(type, setup)
      script = <<~PY
        import sys, datetime
        from pptx.chart.data import CategoryChartData
        from pptx.enum.chart import XL_CHART_TYPE as XL
        cd = CategoryChartData()
        #{setup}
        sys.stdout.write(cd.xml_bytes(XL.#{type}).decode())
      PY
      out, err, status = Open3.capture3("python3", "-c", script)
      raise "xml oracle failed: #{err}" unless status.success?

      out
    end

    def ours(type)
      data = Pptx::ChartData.new
      yield data
      Pptx::ChartXmlWriter.write(Pptx::Enum::XL_CHART_TYPE.fetch(type), data)
    end

    GROUPED_PY = <<~PY
      east = cd.categories.add_category("East")
      north = east.add_sub_category("North")
      north.add_sub_category("A")
      north.add_sub_category("B")
      east.add_sub_category("South").add_sub_category("C")
      cd.categories.add_category("West").add_sub_category("Coast").add_sub_category("D & E")
      cd.add_series("Q1", (1, 2, 3, 4))
      cd.add_series("Q2", (5, None, 7, 8.5))
    PY

    def grouped(data)
      data.categories = { "East" => { "North" => %w[A B], "South" => %w[C] },
                          "West" => { "Coast" => ["D & E"] } }
      data.add_series("Q1", [1, 2, 3, 4])
      data.add_series("Q2", [5, nil, 7, 8.5])
    end

    %i[COLUMN_CLUSTERED LINE AREA PIE RADAR].each do |type|
      it "writes three levels of categories as python-pptx does for #{type}" do
        expect(ours(type) { |data| grouped(data) }).to eq(python_xml(type, GROUPED_PY))
      end
    end

    DATED_PY = <<~PY
      cd.categories = [datetime.date(2024, 1, 15), datetime.date(2024, 2, 1), datetime.date(1900, 2, 28)]
      cd.add_series("Q1", (1, 2, 3))
    PY

    def dated(data)
      data.categories = [Date.new(2024, 1, 15), Date.new(2024, 2, 1), Date.new(1900, 2, 28)]
      data.add_series("Q1", [1, 2, 3])
    end

    # Bar, line and area draw a date axis for date categories; others keep
    # a category axis but still cache the dates as serial numbers.
    %i[COLUMN_CLUSTERED BAR_STACKED LINE_MARKERS AREA_STACKED RADAR PIE].each do |type|
      it "writes date categories as python-pptx does for #{type}" do
        expect(ours(type) { |data| dated(data) }).to eq(python_xml(type, DATED_PY))
      end
    end

    # Only a flat list can be numeric: grouped categories are always text,
    # even when the groups are numbers such as years.
    it "writes numerically-labelled groups as text, as python-pptx does" do
      setup = <<~PY
        for year, quarters in ((2023, ("Q3", "Q4")), (2024, ("Q1",))):
            group = cd.categories.add_category(year)
            for q in quarters:
                group.add_sub_category(q)
        cd.add_series("S", (1, 2, 3))
      PY
      xml = ours(:COLUMN_CLUSTERED) do |data|
        data.categories = { 2023 => %w[Q3 Q4], 2024 => %w[Q1] }
        data.add_series("S", [1, 2, 3])
      end
      expect(xml).to eq(python_xml("COLUMN_CLUSTERED", setup))
    end

    it "writes a custom date format as python-pptx does" do
      setup = "#{DATED_PY}cd.categories.number_format = 'mmm yy'\n"
      xml = ours(:LINE) do |data|
        dated(data)
        data.categories.number_format = "mmm yy"
      end
      expect(xml).to eq(python_xml("LINE", setup))
    end

    # The category cache takes its format from the categories, not from the
    # series' number format.
    it "formats numeric categories by the categories, not the series" do
      setup = <<~PY
        cd = CategoryChartData(number_format="0.0%")
        cd.categories = (1, 2.5)
        cd.add_series("S", (3, 4))
      PY
      data = Pptx::ChartData.new(number_format: "0.0%")
      data.categories = [1, 2.5]
      data.add_series("S", [3, 4])
      xml = Pptx::ChartXmlWriter.write(Pptx::Enum::XL_CHART_TYPE::COLUMN_CLUSTERED, data)
      expect(xml).to eq(python_xml("COLUMN_CLUSTERED", setup))
    end

    it "writes a series shorter than the categories as python-pptx does" do
      setup = "cd.categories = ('a', 'b', 'c')\ncd.add_series('S', (1, 2))\n"
      xml = ours(:COLUMN_CLUSTERED) do |data|
        data.categories = %w[a b c]
        data.add_series("S", [1, 2])
      end
      expect(xml).to eq(python_xml("COLUMN_CLUSTERED", setup))
    end

    def chart_oracle(path)
      require_oracle!("openpyxl", available: Pptx::Spec::Differential.openpyxl_available?)
      out, err, status = Open3.capture3("python3", File.expand_path("../../tools/chart_oracle.py", __dir__),
                                        path)
      raise "chart oracle failed: #{err}" unless status.success?

      JSON.parse(out)["charts"].first
    end

    def saved(type)
      Tempfile.create(["categories", ".pptx"]) do |file|
        file.close
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        data = Pptx::ChartData.new
        yield data
        slide.shapes.add_chart(type, data, at: [0, 0], size: [Pptx.inches(6), Pptx.inches(4)])
        prs.save(file.path)
        chart_oracle(file.path)
      end
    end

    # The workbook is written without XlsxWriter, so it is checked by what a
    # spreadsheet reader finds in it.
    it "lays grouped categories out in the workbook as python-pptx does" do
      chart = saved(:column_clustered) { |data| grouped(data) }
      expect(chart["grid"]).to eq(
        [[nil, nil, nil, "Q1", "Q2"],
         ["East", "North", "A", 1, 5],
         [nil, nil, "B", 2, nil],
         [nil, "South", "C", 3, 7],
         ["West", "Coast", "D & E", 4, 8.5]]
      )
    end

    it "writes date categories to the workbook as dates" do
      chart = saved(:line) { |data| dated(data) }
      expect(chart["grid"].drop(1).map(&:first))
        .to eq(["2024-01-15T00:00:00", "2024-02-01T00:00:00", "1900-02-28T00:00:00"])
    end

    it "is read back by python-pptx as the categories it was given" do
      chart = saved(:column_clustered) { |data| grouped(data) }
      expect(chart["categories"]).to eq(["A", "B", "C", "D & E"])
    end
  end
end
