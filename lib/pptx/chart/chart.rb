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

    # Replace this chart's categories and series with those in +chart_data+.
    #
    # Series-level formatting is left alone. If there are more series than
    # before, the extra ones are cloned from the last series of the last plot
    # so they inherit its formatting; if fewer, the surplus is removed along
    # with any plot left empty.
    def replace_data(chart_data)
      SeriesRewriter.new(chart_data).rewrite(@element)
      part.workbook.replace_with(chart_data.xlsx_blob)
      self
    end

    # Whether a legend is drawn.
    def legend? = !chart_element.legend.nil?

    def legend=(value)
      value ? chart_element.get_or_add_legend : chart_element.remove_legend
      @legend = nil
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
        return
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
  end

  # One series as read back from a chart's cached XML.
  class ChartSeriesView < ElementProxy
    def initialize(ser, chart)
      super(ser)
      @chart = chart
    end

    # The fill and outline of this series.
    def format = @format ||= ChartFormat.new(@element)

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

  # Rewrites a chart's series data in place, leaving formatting alone.
  class SeriesRewriter
    def initialize(chart_data)
      @chart_data = chart_data
    end

    def rewrite(chart_space)
      plot_area = chart_space.chart.plotArea
      adjust_series_count(plot_area, @chart_data.series.size)
      plot_area.series_elements.zip(@chart_data.series) do |ser, series_data|
        rewrite_series(ser, series_data)
      end
      self
    end

    private

    def adjust_series_count(plot_area, wanted)
      difference = wanted - plot_area.series_elements.size
      if difference.positive?
        clone_series(plot_area, difference)
      elsif difference.negative?
        trim_series(plot_area, -difference)
      end
    end

    # New series are copied from the last one so they pick up its formatting;
    # only their position in the chart is changed.
    def clone_series(plot_area, count)
      last = plot_area.series_elements.last
      raise Error, "cannot add series to a chart that has none" if last.nil?

      next_index = plot_area.series_elements.size
      count.times do |offset|
        copy = last.parent.build_from_xml(last.node.to_xml)
        copy.get_or_add_idx.val = next_index + offset
        copy.get_or_add_order.val = next_index + offset
        last.node.add_next_sibling(copy.node)
        last = copy
      end
    end

    # A plot left with no series is removed too; an empty plot element is not
    # valid, and PowerPoint would have nothing to draw for it.
    def trim_series(plot_area, count)
      plot_area.series_elements.last(count).each { |ser| ser.parent.remove(ser) }
      plot_area.plot_elements.each do |plot|
        plot_area.remove(plot) if plot.ser_list.empty?
      end
    end

    def rewrite_series(ser, series_data)
      ser.remove_tx
      ser.remove_cat
      ser.remove_val
      tx, cat, val = SeriesFragments.new(@chart_data, series_data).fragments
      ser.insert_tx(ser.build_from_xml(tx))
      ser.insert_cat(ser.build_from_xml(cat))
      ser.insert_val(ser.build_from_xml(val))
    end
  end

  # Builds the `c:tx`, `c:cat` and `c:val` of one series as standalone XML.
  #
  # It borrows the writer that produces those fragments when a chart is
  # created, so a rewritten series and a freshly written one cannot drift
  # apart.
  class SeriesFragments < ChartXmlWriter::Base
    def initialize(chart_data, series)
      super(nil, chart_data)
      @series = series
    end

    # @return [Array(String, String, String)] the tx, cat and val fragments
    def fragments
      [namespaced(tx_xml(@series)), namespaced(cat_xml(@series)), namespaced(val_xml(@series))]
    end

    private

    # The fragments are written for a document that already declares the chart
    # namespaces; parsed on their own they need their own declaration.
    def namespaced(fragment)
      fragment.sub(/\A<(c:\w+)/, %(<\\1 #{Oxml::Ns.nsdecls("c", "a", "r")}))
    end
  end
end
