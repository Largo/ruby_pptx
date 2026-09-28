# frozen_string_literal: true

require "ruby_pptx/pattern_matching"

require "ruby_pptx/element_proxy"
require "ruby_pptx/enum/chart"
require "ruby_pptx/chart/format"

module Pptx
  # A chart, reached through the graphic frame that holds it.
  class Chart < PartElementProxy
    include PatternMatching

    pattern_keys :chart_type, :plot_type, :categories, :series

    PLOT_TAG_TO_TYPE = {
      "c:barChart" => :BAR, "c:lineChart" => :LINE, "c:pieChart" => :PIE,
      "c:doughnutChart" => :DOUGHNUT, "c:areaChart" => :AREA, "c:area3DChart" => :AREA,
      "c:radarChart" => :RADAR, "c:scatterChart" => :XY, "c:bubbleChart" => :BUBBLE
    }.freeze

    # The family of plot this chart draws with, e.g. :BAR. {#chart_type} is
    # the precise type.
    def plot_type
      element = @element.plot_element
      element && PLOT_TAG_TO_TYPE[element.nsptag]
    end

    # The chart type of the first plot, e.g. COLUMN_STACKED. Recovered from
    # the plot's direction, grouping, markers and other settings, as
    # python-pptx's PlotTypeInspector does.
    #
    # @return [Pptx::Enum::Member] a member of XL_CHART_TYPE
    def chart_type = ChartTypeInspector.chart_type(plot_area.plot_elements.first)

    # The chart style number, 1 to 48, or nil when none is set.
    def chart_style = @element.style&.val

    def chart_style=(value)
      @element.remove_style
      @element.get_or_add_style.val = value unless value.nil?
    end

    # The default font for all text in the chart, created on first use.
    def font = @font ||= Font.new(@element.defRPr)

    # Every series in the chart: plot by plot, and within a plot in the order
    # it draws them.
    #
    # @return [Array<ChartSeriesView>]
    def series
      plot_area.plot_elements.flat_map { |plot| plot.sers.map { |ser| ChartSeriesView.for(ser, self) } }
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

    # A String sets the title's text; true shows a title with no text of its
    # own, which PowerPoint fills in from the series; nil or false removes it.
    #
    # Removing it also records that it was deleted. Without that flag a
    # single-series chart grows its generated title back when opened.
    def title=(value)
      case value
      when nil, false
        chart_element.remove_title
        chart_element.get_or_add_autoTitleDeleted.val = true
      when true then chart_element.get_or_add_title
      else ChartTitle.new(chart_element.get_or_add_title).text = value
      end
    end

    # The plots -- "chart groups" in the MS API -- this chart draws.
    #
    # Nearly every chart has exactly one; a combo chart has several, which is
    # why this is a collection rather than a property of the chart.
    def plots = plot_area.plot_elements.map { |element| ChartPlot.new(element, self) }

    # The horizontal axis: a category or date axis, or, on an XY or bubble
    # chart, the first value axis. nil for a chart with no axes, such as a pie.
    #
    # @return [CategoryAxis, DateAxis, ValueAxis, nil]
    def category_axis
      if (element = plot_area.find("c:catAx")) then CategoryAxis.new(element)
      elsif (element = plot_area.find("c:dateAx")) then DateAxis.new(element)
      elsif (element = plot_area.value_axes.first) then ValueAxis.new(element)
      end
    end

    # The vertical value axis, or nil for a chart with none.
    #
    # An XY or bubble chart has two value axes, and this is the second -- the
    # vertical one -- as in python-pptx. {#category_axis} is the horizontal.
    def value_axis
      axes = plot_area.value_axes
      return nil if axes.empty?

      ValueAxis.new(axes[axes.size > 1 ? 1 : 0])
    end

    def value_axes = plot_area.value_axes.map { |element| ValueAxis.new(element) }

    # The category labels cached in the chart XML.
    def categories
      first = @element.series_elements.first
      return [] if first.nil?

      first.xpath("./c:cat//c:pt/c:v").map(&:text)
    end

    def inspect = "#<Pptx::Chart #{chart_type.name} series=#{series.size}>"

    private

    def chart_element = @element.chart

    def plot_area = chart_element.plotArea
  end

  # Works out the full chart type from a plot element, after python-pptx's
  # PlotTypeInspector. A bar and a column chart, or a clustered and a stacked
  # one, share an element and differ only in its settings.
  module ChartTypeInspector
    XL = Enum::XL_CHART_TYPE

    GROUPED = {
      "c:areaChart" => { "standard" => :AREA, "stacked" => :AREA_STACKED,
                         "percentStacked" => :AREA_STACKED_100 },
      "c:area3DChart" => { "standard" => :THREE_D_AREA, "stacked" => :THREE_D_AREA_STACKED,
                           "percentStacked" => :THREE_D_AREA_STACKED_100 }
    }.freeze

    BAR = {
      "bar" => { "clustered" => :BAR_CLUSTERED, "stacked" => :BAR_STACKED,
                 "percentStacked" => :BAR_STACKED_100 },
      "col" => { "clustered" => :COLUMN_CLUSTERED, "stacked" => :COLUMN_STACKED,
                 "percentStacked" => :COLUMN_STACKED_100 }
    }.freeze

    LINE = {
      true => { "standard" => :LINE_MARKERS, "stacked" => :LINE_MARKERS_STACKED,
                "percentStacked" => :LINE_MARKERS_STACKED_100 },
      false => { "standard" => :LINE, "stacked" => :LINE_STACKED,
                 "percentStacked" => :LINE_STACKED_100 }
    }.freeze

    module_function

    # @return [Pptx::Enum::Member]
    def chart_type(plot)
      XL.fetch(name_for(plot))
    end

    # How to tell the types apart, per plot element.
    RULES = {
      "c:areaChart" => ->(plot) { GROUPED.fetch("c:areaChart").fetch(plot.grouping_val) },
      "c:area3DChart" => ->(plot) { GROUPED.fetch("c:area3DChart").fetch(plot.grouping_val) },
      "c:barChart" => ->(plot) { BAR.fetch(plot.barDir.val).fetch(plot.grouping_val) },
      "c:lineChart" => ->(plot) { LINE.fetch(line_markers?(plot)).fetch(plot.grouping_val) },
      "c:bubbleChart" => ->(plot) { bubble(plot) },
      "c:doughnutChart" => ->(plot) { exploded?(plot) ? :DOUGHNUT_EXPLODED : :DOUGHNUT },
      "c:pieChart" => ->(plot) { exploded?(plot) ? :PIE_EXPLODED : :PIE },
      "c:radarChart" => ->(plot) { radar(plot) },
      "c:scatterChart" => ->(plot) { scatter(plot) }
    }.freeze

    def name_for(plot)
      rule = RULES.fetch(plot.nsptag) { raise Error, "no chart type for #{plot.nsptag}" }
      rule.call(plot)
    end

    def line_markers?(plot) = plot.xpath('c:ser/c:marker/c:symbol[@val="none"]').empty?

    def exploded?(plot) = !plot.xpath("./c:ser/c:explosion").empty?

    def first_symbol(plot) = plot.xpath("c:ser/c:marker/c:symbol").first&.get("val")

    def bubble(plot)
      bubble3d = plot.xpath("c:ser/c:bubble3D").first
      bubble3d&.val ? :BUBBLE_THREE_D_EFFECT : :BUBBLE
    end

    def radar(plot)
      style = plot.xpath("c:radarStyle").first&.get("val")
      return :RADAR if style.nil?
      return :RADAR_FILLED if style == "filled"

      first_symbol(plot) == "none" ? :RADAR : :RADAR_MARKERS
    end

    def scatter(plot)
      no_markers = first_symbol(plot) == "none"
      case plot.xpath("c:scatterStyle").first&.get("val")
      when "lineMarker"
        return :XY_SCATTER unless plot.xpath("c:ser/c:spPr/a:ln/a:noFill").empty?

        no_markers ? :XY_SCATTER_LINES_NO_MARKERS : :XY_SCATTER_LINES
      when "smoothMarker"
        no_markers ? :XY_SCATTER_SMOOTH_NO_MARKERS : :XY_SCATTER_SMOOTH
      else :XY_SCATTER
      end
    end
  end

  # One series as read back from a chart's cached XML.
  #
  # {.for} returns the subclass matching the plot the series belongs to, so a
  # series offers only what its kind of chart supports: {LineSeriesView#smooth},
  # {BarSeriesView#invert_if_negative}, markers on the kinds that draw them.
  class ChartSeriesView < ElementProxy
    include PatternMatching

    pattern_keys :name, :values

    # @return [ChartSeriesView] the subclass for the series' plot
    def self.for(ser, chart)
      klass = {
        "c:areaChart" => AreaSeriesView, "c:area3DChart" => AreaSeriesView, "c:barChart" => BarSeriesView,
        "c:bubbleChart" => BubbleSeriesView, "c:doughnutChart" => PieSeriesView,
        "c:lineChart" => LineSeriesView, "c:pieChart" => PieSeriesView, "c:radarChart" => RadarSeriesView,
        "c:scatterChart" => XySeriesView
      }.fetch(ser.parent.nsptag, ChartSeriesView)
      klass.new(ser, chart)
    end

    def initialize(ser, chart)
      super(ser)
      @chart = chart
    end

    # The fill and outline of this series.
    def format = @format ||= ChartFormat.new(@element)

    # The series' chart-wide index.
    def index = @element.idx.val

    def name = @element.xpath("./c:tx//c:pt/c:v").first&.text.to_s

    # Cached values, with nil where the chart records a gap.
    def values = cached(@element.val)

    def inspect = "#<#{self.class.name} #{name.inspect}>"

    private

    def cached(source)
      return [] if source.nil?

      count = source.xpath(".//c:ptCount/@val").first&.value.to_i
      by_index = source.xpath(".//c:pt").to_h { |pt| [pt.idx, pt.value] }
      Array.new(count) { |i| by_index[i] }
    end
  end

  # Data labels and per-point formatting, for series plotted by category.
  module CategorySeriesFeatures
    # This series' own data labels, created on first use.
    def data_labels = @data_labels ||= ChartDataLabels.new(@element.get_or_add_dLbls)

    # The points of this series, one per category.
    def points = ChartPoints.new(@element, @element.cat_ptCount_val)
  end

  # Markers, for the kinds of series that draw them.
  module MarkerFeatures
    def marker = @marker ||= Marker.new(@element)
  end

  class AreaSeriesView < ChartSeriesView
    include CategorySeriesFeatures
  end

  class PieSeriesView < ChartSeriesView
    include CategorySeriesFeatures
  end

  class BarSeriesView < ChartSeriesView
    include CategorySeriesFeatures

    # Whether a negative bar is drawn in inverted colours. An absent
    # `c:invertIfNegative` reads as true, the schema default.
    def invert_if_negative? = @element.invertIfNegative.nil? || @element.invertIfNegative.val

    def invert_if_negative=(value)
      @element.get_or_add_invertIfNegative.val = value ? true : false
    end
  end

  class LineSeriesView < ChartSeriesView
    include CategorySeriesFeatures
    include MarkerFeatures

    # Whether the line is drawn as a smooth curve. An absent `c:smooth` reads
    # as true, the schema default.
    def smooth? = @element.smooth.nil? || @element.smooth.val

    def smooth=(value)
      @element.get_or_add_smooth.val = value ? true : false
    end
  end

  class RadarSeriesView < ChartSeriesView
    include CategorySeriesFeatures
    include MarkerFeatures
  end

  # A scatter series, whose values are its y values.
  class XySeriesView < ChartSeriesView
    include MarkerFeatures

    def values = cached(@element.yVal)

    def points = ChartPoints.new(@element, [@element.xVal_ptCount_val, @element.yVal_ptCount_val].min)
  end

  class BubbleSeriesView < XySeriesView
    def points
      counts = [@element.xVal_ptCount_val, @element.yVal_ptCount_val, @element.bubbleSize_ptCount_val]
      ChartPoints.new(@element, counts.min)
    end
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
