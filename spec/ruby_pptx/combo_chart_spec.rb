# frozen_string_literal: true

require "json"
require "open3"

# A combo chart draws more than one plot over the same categories. python-pptx
# can read such a chart but only ever writes one plot, so there is no
# differential oracle. These check the structure against the OOXML schema's
# requirements and confirm python-pptx reads the result correctly.
RSpec.describe Pptx::ComboChartBuilder do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }

  let(:data) do
    Pptx::ChartData.new.tap do |d|
      d.categories = %w[East West Mid]
      d.add_series("Revenue", [100, 120, 90])
      d.add_series("Margin", [0.12, 0.18, 0.09])
    end
  end

  def combo(secondary: true, &block)
    block ||= lambda { |c|
      c.plot :column_clustered, series: "Revenue"
      c.plot :line, series: "Margin", secondary_axis: secondary
    }
    slide.shapes.add_combo_chart(data, at: [Pptx.inches(1), Pptx.inches(1)],
                                       size: [Pptx.inches(8), Pptx.inches(5)], &block)
  end

  describe "selecting series" do
    subject(:builder) { described_class.new(data) }

    it "accepts a series by name, by index, or as a list" do
      aggregate_failures do
        expect(builder.plot(:line, series: "Margin").specs.last.series.map(&:name))
          .to eq(["Margin"])
        expect(builder.plot(:line, series: 0).specs.last.series.map(&:name))
          .to eq(["Revenue"])
        expect(builder.plot(:line, series: ["Revenue", 1]).specs.last.series.map(&:name))
          .to eq(%w[Revenue Margin])
      end
    end

    it "draws every series when none is named" do
      expect(builder.plot(:line).specs.last.series.size).to eq(2)
    end

    it "raises for a series that does not exist" do
      expect { builder.plot(:line, series: "Nope") }
        .to raise_error(Pptx::NotFoundError, /no series/)
    end

    # A pie has no category axis to share, so it cannot be combined.
    it "refuses a chart type that cannot share a category axis" do
      aggregate_failures do
        expect { builder.plot(:pie) }.to raise_error(Pptx::Error, /cannot share a category axis/)
        expect { builder.plot(:xy_scatter) }.to raise_error(Pptx::Error, /cannot share/)
        expect { builder.plot(:doughnut) }.to raise_error(Pptx::Error, /cannot share/)
      end
    end

    it "allows bar, line and area together" do
      expect do
        builder.plot(:column_clustered, series: 0)
        builder.plot(:line, series: 1)
      end.not_to raise_error
    end
  end

  describe "validation" do
    subject(:builder) { described_class.new(data) }

    it "needs at least two plots" do
      builder.plot(:line)
      expect { builder.validate! }.to raise_error(Pptx::Error, /at least two plots/)
    end

    # A series drawn twice would appear twice in the legend and confuse the
    # chart-wide series indices.
    it "refuses to draw the same series in two plots" do
      builder.plot(:column_clustered, series: "Revenue")
      builder.plot(:line, series: "Revenue")
      expect { builder.validate! }.to raise_error(Pptx::Error, /more than one plot/)
    end
  end

  describe "the XML it writes" do
    subject(:chart) { combo.chart }

    def plot_area_children(element)
      element.xpath("//c:chart/c:plotArea/*").map(&:nsptag)
    end

    # CT_PlotArea is layout?, then one or more chart groups, then the axes.
    it "orders the plot area as the schema requires" do
      expect(plot_area_children(chart.element))
        .to eq(["c:layout", "c:barChart", "c:lineChart", "c:catAx", "c:valAx",
                "c:catAx", "c:valAx"])
    end

    it "writes one plot element per plot" do
      expect(chart.plots.map { |p| p.element.nsptag }).to eq(["c:barChart", "c:lineChart"])
    end

    # Each plot must name exactly two axes, and every id must be backed by a
    # real axis element: a dangling reference is invalid and PowerPoint
    # rejects the file.
    it "gives every plot exactly two axis ids" do
      counts = chart.element.xpath("//c:chart/c:plotArea/*")
                    .select { |e| e.nsptag.end_with?("Chart") }
                    .map { |e| e.xpath("./c:axId").size }
      expect(counts).to eq([2, 2])
    end

    it "backs every referenced axis id with an axis, and uses every axis" do
      referenced = chart.element.xpath("//c:chart/c:plotArea/*/c:axId/@val").map(&:value).uniq
      defined_ids = chart.element.xpath("//c:chart/c:plotArea/c:catAx/c:axId/@val").map(&:value) +
                    chart.element.xpath("//c:chart/c:plotArea/c:valAx/c:axId/@val").map(&:value)
      aggregate_failures do
        expect(referenced - defined_ids).to be_empty
        expect(defined_ids - referenced).to be_empty
      end
    end

    it "puts each plot's series in its own plot element" do
      bar, line = chart.plots.map { |p| p.element.ser_list }
      aggregate_failures do
        expect(bar.size).to eq(1)
        expect(line.size).to eq(1)
        expect(bar.first.xpath("./c:tx//c:v").first.text).to eq("Revenue")
        expect(line.first.xpath("./c:tx//c:v").first.text).to eq("Margin")
      end
    end

    # Series indices are chart-wide, not per-plot, so they must stay distinct
    # across the whole chart.
    it "numbers series across the chart rather than within each plot" do
      indices = chart.element.xpath("//c:ser/c:idx/@val").map { |a| a.value.to_i }
      expect(indices).to eq([0, 1])
    end
  end

  describe "the secondary axis" do
    it "adds a value axis on the right, crossing at the maximum" do
      chart = combo.chart
      secondary = chart.element.xpath("//c:chart/c:plotArea/c:valAx").last
      aggregate_failures do
        expect(secondary.xpath("./c:axPos/@val").first.value).to eq("r")
        expect(secondary.xpath("./c:crosses/@val").first.value).to eq("max")
      end
    end

    # The schema pairs every value axis with a category axis, but the
    # categories are already drawn by the primary pair, so the second one is
    # hidden.
    it "hides the category axis that comes with it" do
      chart = combo.chart
      secondary_cat = chart.element.xpath("//c:chart/c:plotArea/c:catAx").last
      expect(secondary_cat.xpath("./c:delete/@val").first.value).to eq("1")
    end

    it "writes only one pair of axes when no plot asks for a second" do
      chart = combo(secondary: false).chart
      aggregate_failures do
        expect(chart.element.xpath("//c:chart/c:plotArea/c:valAx").size).to eq(1)
        expect(chart.element.xpath("//c:chart/c:plotArea/c:catAx").size).to eq(1)
      end
    end
  end

  describe "the embedded workbook" do
    # One worksheet holds every series, whichever plot draws it, because the
    # references are computed from a single data layout.
    it "holds all the series in one sheet" do
      combo
      partnames = presentation.part.package.parts.map { |p| p.partname.to_s }
      aggregate_failures do
        expect(partnames.grep(%r{/ppt/embeddings/}).size).to eq(1)
        expect(data.series[1].values_ref).to eq("Sheet1!$C$2:$C$4")
      end
    end
  end

  describe "what python-pptx sees in our output" do
    def oracle_path = File.expand_path("../../tools/combo_chart_oracle.py", __dir__)

    before { require_oracle! }

    # python-pptx understands multi-plot charts even though it never writes
    # one, so it makes a good reader-of-record for the structure.
    it "reads back as two plots with the right series and categories" do
      combo
      result = Tempfile.create(["combo", ".pptx"]) do |file|
        file.close
        presentation.save(file.path)
        out, err, status = Open3.capture3("python3", oracle_path, file.path)
        raise "combo oracle failed: #{err}" unless status.success?

        JSON.parse(out)
      end

      chart = result["charts"].first
      aggregate_failures do
        expect(chart["plots"].map { |p| p["kind"] }).to eq(%w[BarPlot LinePlot])
        expect(chart["plots"][0]["series"]).to eq(["Revenue"])
        expect(chart["plots"][1]["series"]).to eq(["Margin"])
        expect(chart["plots"].map { |p| p["categories"] })
          .to all(eq(%w[East West Mid]))
        expect(chart["value_axis_readable"]).to be(true)
      end
    end
  end
end
