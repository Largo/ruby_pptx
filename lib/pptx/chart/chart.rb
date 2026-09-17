# frozen_string_literal: true

require "pptx/element_proxy"
require "pptx/enum/chart"

module Pptx
  # A chart, reached through the graphic frame that holds it.
  #
  # Creating a chart is the well-covered path; reading one back is currently
  # limited to its type, series and cached values.
  class Chart < PartElementProxy
    PLOT_TAG_TO_TYPE = {
      "c:barChart" => :BAR, "c:lineChart" => :LINE, "c:pieChart" => :PIE,
      "c:doughnutChart" => :DOUGHNUT, "c:areaChart" => :AREA,
      "c:radarChart" => :RADAR, "c:scatterChart" => :XY, "c:bubbleChart" => :BUBBLE
    }.freeze

    # The family of plot this chart draws with, e.g. :BAR.
    #
    # This is the plot element rather than the full XL_CHART_TYPE, which cannot
    # always be recovered from the XML: a clustered and a stacked bar chart
    # differ only by their grouping.
    def plot_type
      element = @element.plot_element
      element && PLOT_TAG_TO_TYPE[element.nsptag]
    end

    # @return [Array<ChartSeriesView>] the series, in plot order
    def series
      @element.series_elements.map { |ser| ChartSeriesView.new(ser, self) }
    end

    def legend? = !@element.xpath("./c:chart/c:legend").empty?

    # The category labels cached in the chart XML.
    def categories
      first = @element.series_elements.first
      return [] if first.nil?

      first.xpath("./c:cat//c:pt/c:v").map(&:text)
    end

    def inspect = "#<Pptx::Chart #{plot_type} series=#{series.size}>"
  end

  # One series as read back from a chart's cached XML.
  class ChartSeriesView < ElementProxy
    def initialize(ser, chart)
      super(ser)
      @chart = chart
    end

    def name = @element.xpath("./c:tx//c:pt/c:v").first&.text.to_s

    # Cached values, with nil where the chart records a gap.
    def values
      points = @element.xpath("./c:val//c:pt")
      count = @element.xpath("./c:val//c:ptCount/@val").first&.value.to_i
      by_index = points.to_h do |pt|
        [pt.get("idx").to_i, Float(pt.xpath("./c:v").first.text)]
      end
      Array.new(count) { |i| by_index[i] }
    end

    def inspect = "#<Pptx::ChartSeriesView #{name.inspect}>"
  end
end
