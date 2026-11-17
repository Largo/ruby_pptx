# frozen_string_literal: true

require "tmpdir"

# Regression specs for Chart#replace_data on scatter, bubble and combo charts
# (GitHub issues #1 and #2).
RSpec.describe "Chart#replace_data on scatter, bubble and combo charts" do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }
  let(:at) { [Pptx.inches(1), Pptx.inches(1)] }
  let(:size) { [Pptx.inches(5), Pptx.inches(3)] }

  def category_data(categories, series)
    Pptx::ChartData.new.tap do |d|
      d.categories = categories
      series.each { |name, values| d.add_series(name, values) }
    end
  end

  def xy_data(series, klass: Pptx::XyChartData)
    klass.new.tap do |d|
      series.each { |name, points| d.add_series(name, points: points) }
    end
  end

  # Save and open again, so the assertions see what PowerPoint would read.
  def round_trip(prs)
    Dir.mktmpdir do |dir|
      path = File.join(dir, "out.pptx")
      prs.save(path)
      yield Pptx::Presentation.open(path)
    end
  end

  def reopened_chart(prs)
    prs.slides.first.shapes.find(&:chart?).chart
  end

  # The children of every c:ser must follow CT_Series' sequence, or the part
  # is schema-invalid and PowerPoint offers to repair the file.
  def expect_schema_ordered_series(chart)
    order = Pptx::Oxml::CT_Series::TAG_SEQ
    chart.element.xpath("//c:ser").each do |ser|
      tags = ser.element_children.map(&:nsptag)
      expect(tags.map { |tag| order.index(tag) }).to eq(tags.map { |tag| order.index(tag) }.sort),
                                                     "c:ser children out of order: #{tags.inspect}"
    end
  end

  describe "scatter charts (issue #1)" do
    let(:chart) do
      slide.shapes.add_chart(Pptx::Enum::XL_CHART_TYPE::XY_SCATTER,
                             xy_data([["s1", [[1, 10], [2, 20]]]]), at: at, size: size).chart
    end

    it "rebinds with XyChartData and survives a save" do
      chart.replace_data(xy_data([["s2", [[3, 30], [4, 40], [5, 50]]]]))

      round_trip(presentation) do |prs|
        rebound = reopened_chart(prs)
        ser = rebound.element.xpath("//c:ser").first
        aggregate_failures do
          expect(rebound.series.map(&:name)).to eq(["s2"])
          expect(rebound.series.first.values).to eq([30.0, 40.0, 50.0])
          expect(ser.xpath("./c:xVal//c:pt/c:v").map(&:text)).to eq(%w[3 4 5])
          expect(ser.xpath("./c:cat | ./c:val")).to be_empty
          expect(ser.xpath("./c:xVal").size).to eq(1)
          expect(ser.xpath("./c:yVal").size).to eq(1)
          expect_schema_ordered_series(rebound)
        end
      end
    end

    it "keeps the series formatting (the hidden line of a plain scatter)" do
      chart.replace_data(xy_data([["s2", [[3, 30]]]]))
      expect(chart.element.xpath("//c:ser/c:spPr/a:ln/a:noFill")).not_to be_empty
    end

    it "adds and removes series" do
      chart.replace_data(xy_data([["a", [[1, 1]]], ["b", [[2, 2]]], ["c", [[3, 3]]]]))
      expect(chart.series.map(&:name)).to eq(%w[a b c])
      chart.replace_data(xy_data([["only", [[9, 9]]]]))
      expect(chart.series.map(&:name)).to eq(["only"])
      expect_schema_ordered_series(chart)
    end

    it "refuses category data instead of writing a corrupt c:ser" do
      before = chart.element.xml
      expect { chart.replace_data(category_data(%w[A B], [["s2", [30, 40]]])) }
        .to raise_error(Pptx::Error, /scatter chart.*Pptx::ChartData.*Pptx::XyChartData/)
      expect(chart.element.xml).to eq(before)
    end
  end

  describe "bubble charts" do
    let(:chart) do
      data = xy_data([["b1", [[1, 10, 5], [2, 20, 8]]]], klass: Pptx::BubbleChartData)
      slide.shapes.add_chart(Pptx::Enum::XL_CHART_TYPE::BUBBLE, data, at: at, size: size).chart
    end

    it "rebinds x, y and bubble sizes" do
      chart.replace_data(xy_data([["b2", [[3, 30, 7]]]], klass: Pptx::BubbleChartData))
      ser = chart.element.xpath("//c:ser").first
      aggregate_failures do
        expect(chart.series.first.values).to eq([30.0])
        expect(ser.xpath("./c:bubbleSize//c:pt/c:v").map(&:text)).to eq(%w[7])
        expect(ser.xpath("./c:bubble3D")).not_to be_empty
        expect_schema_ordered_series(chart)
      end
    end

    it "refuses scatter data, which has no bubble sizes" do
      expect { chart.replace_data(xy_data([["x", [[1, 1]]]])) }
        .to raise_error(Pptx::Error, /bubble chart.*Pptx::BubbleChartData/)
    end
  end

  it "refuses XY data on a category chart" do
    chart = slide.shapes.add_chart(:column_clustered, category_data(%w[A], [["s", [1]]]),
                                   at: at, size: size).chart
    expect { chart.replace_data(xy_data([["x", [[1, 1]]]])) }
      .to raise_error(Pptx::Error, /category chart.*Pptx::ChartData/)
  end

  describe "combo charts (issue #2)" do
    let(:chart) do
      data = category_data(%w[Q1 Q2 Q3], [["Revenue", [120, 150, 180]], ["Margin", [0.18, 0.22, 0.25]]])
      slide.shapes.add_combo_chart(data, at: at, size: [Pptx.inches(8), Pptx.inches(5)]) do |combo|
        combo.plot :column_clustered, series: "Revenue"
        combo.plot :line, series: "Margin", secondary_axis: true
      end.chart
    end

    def plot_shape(chart)
      chart.plots.map { |p| [p.element.nsptag, p.series.size] }
    end

    it "rebinds both plots when the series count is unchanged" do
      series = [["Rev2", [100, 110, 120]], ["Mar2", [0.1, 0.2, 0.3]]]
      chart.replace_data(category_data(%w[Jan Feb Mar], series))

      round_trip(presentation) do |prs|
        rebound = reopened_chart(prs)
        aggregate_failures do
          expect(plot_shape(rebound)).to eq([["c:barChart", 1], ["c:lineChart", 1]])
          expect(rebound.series.map(&:name)).to eq(%w[Rev2 Mar2])
          expect(rebound.categories).to eq(%w[Jan Feb Mar])
          expect(rebound.element.xpath("//c:valAx").size).to eq(2)
        end
      end
    end

    it "raises instead of silently dropping the line plot when given fewer series" do
      before = chart.element.xml
      expect { chart.replace_data(category_data(%w[Jan Feb Mar], [["Rev2", [100, 110, 120]]])) }
        .to raise_error(Pptx::Error, /series count of a chart with 2 plots \(2 series, 1 given\)/)
      aggregate_failures do
        expect(plot_shape(chart)).to eq([["c:barChart", 1], ["c:lineChart", 1]])
        expect(chart.element.xml).to eq(before)
      end
    end

    it "raises instead of guessing a plot when given more series" do
      data = category_data(%w[A], [["a", [1]], ["b", [2]], ["c", [3]]])
      expect { chart.replace_data(data) }
        .to raise_error(Pptx::Error, /2 series, 3 given/)
      expect(plot_shape(chart)).to eq([["c:barChart", 1], ["c:lineChart", 1]])
    end
  end
end
