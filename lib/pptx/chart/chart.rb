# frozen_string_literal: true

require "pptx/element_proxy"
require "pptx/enum/chart"
require "pptx/chart/format"

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

    # Whether a legend is drawn.
    def legend? = !chart_element.legend.nil?

    def legend=(value)
      value ? chart_element.get_or_add_legend : chart_element.remove_legend
      @legend = nil
      value
    end

    # @return [ChartLegend, nil] nil when the chart has no legend
    def legend
      element = chart_element.legend
      element.nil? ? nil : (@legend ||= ChartLegend.new(element))
    end

    def title? = !chart_element.title.nil?

    # @return [ChartTitle, nil]
    def title
      element = chart_element.title
      element.nil? ? nil : ChartTitle.new(element)
    end

    # Give the chart a title, or remove it with nil.
    def title=(text)
      if text.nil?
        chart_element.remove_title
        return nil
      end

      ChartTitle.new(ensure_title).text = text
    end

    # The plots -- "chart groups" in the MS API -- this chart draws.
    #
    # Nearly every chart has exactly one; a combo chart has several, which is
    # why this is a collection rather than a property of the chart.
    def plots = plot_area.plot_elements.map { |element| ChartPlot.new(element, self) }

    # The category axis, or nil for a chart type that has none, such as a pie.
    def category_axis
      element = plot_area.category_axis
      element.nil? ? nil : ChartAxis.new(element)
    end

    # The value axis, or nil for a chart type that has none.
    #
    # A scatter or bubble chart has two value axes; this returns the first,
    # which is the horizontal one. {#value_axes} gives both.
    def value_axis
      element = plot_area.value_axis
      element.nil? ? nil : ChartAxis.new(element)
    end

    def value_axes = plot_area.value_axes.map { |element| ChartAxis.new(element) }

    # The category labels cached in the chart XML.
    def categories
      first = @element.series_elements.first
      return [] if first.nil?

      first.xpath("./c:cat//c:pt/c:v").map(&:text)
    end

    def inspect = "#<Pptx::Chart #{plot_type} series=#{series.size}>"

    private

    def chart_element = @element.chart

    def plot_area = chart_element.plotArea

    def ensure_title
      chart_element.title || begin
        created = Oxml::CT_Title.new_title(chart_element)
        chart_element.insert_title(created)
        created
      end
    end

    def chart_element = @element.chart

    def plot_area = chart_element.plotArea
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
