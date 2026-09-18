# frozen_string_literal: true

require "ruby_pptx/chart/xml_writer"

module Pptx
  # Describes a chart that draws more than one plot over the same categories --
  # a bar chart with a line over it, most often.
  #
  # python-pptx can read such a chart but only ever writes one plot, so this is
  # new ground rather than a port.
  #
  #   slide.shapes.add_combo_chart(data, at: [x, y], size: [w, h]) do |combo|
  #     combo.plot :column_clustered, series: "Revenue"
  #     combo.plot :line, series: "Margin", secondary_axis: true
  #   end
  #
  # All the series live in one {ChartData}, because they share a worksheet;
  # each plot names the ones it draws.
  class ComboChartBuilder
    # What a plot draws and how.
    PlotSpec = Data.define(:chart_type, :series, :secondary_axis)

    # The families that can share a category axis. A pie or a scatter cannot,
    # so they are not combinable.
    COMBINABLE = %i[bar line area].freeze

    attr_reader :chart_data, :specs

    def initialize(chart_data)
      @chart_data = chart_data
      @specs = []
    end

    # Add a plot drawing +series+.
    #
    # @param series [String, Integer, Array] series names or indices; all of
    #   them when omitted
    # @param secondary_axis [Boolean] draw against a second value axis on the
    #   right, for values on a different scale
    def plot(chart_type, series: nil, secondary_axis: false)
      member = Enum::XL_CHART_TYPE.fetch(chart_type)
      family = ChartXmlWriter.family_of(member)
      unless COMBINABLE.include?(family)
        raise Error,
              "a #{member.name} chart cannot share a category axis, so it " \
              "cannot be part of a combo chart"
      end

      @specs << PlotSpec.new(chart_type: member, series: resolve(series),
                             secondary_axis: secondary_axis)
      self
    end

    # @raise [Error] when the combination could not be drawn
    def validate!
      raise Error, "a combo chart needs at least two plots" if @specs.size < 2

      drawn = @specs.flat_map(&:series)
      duplicated = drawn.map(&:index).tally.select { |_, count| count > 1 }.keys
      raise Error, "series #{duplicated.inspect} appear in more than one plot" unless duplicated.empty?

      self
    end

    def secondary_axis? = @specs.any?(&:secondary_axis)

    private

    # Accepts a name, an index, or a list of either; nil means every series.
    def resolve(selector)
      return @chart_data.series if selector.nil?

      Array(selector).map do |key|
        case key
        when Integer then @chart_data.series[key]
        when String then @chart_data.series.find { |s| s.name == key }
        else key
        end || raise(NotFoundError, "no series #{key.inspect} in this chart data")
      end
    end
  end

  # Writes the `c:chartSpace` for a combo chart.
  #
  # The plot area holds one element per plot followed by the axes, which is the
  # order the schema requires. Every plot names exactly two axis ids and each
  # must be backed by a real axis element, so a secondary value axis brings a
  # category axis with it -- a hidden one, since the categories are already
  # drawn by the primary pair.
  class ComboChartWriter < ChartXmlWriter::Base
    CAT_AX_ID = "-2068027336"
    VAL_AX_ID = "-2113994440"
    SECONDARY_CAT_AX_ID = "-2068027337"
    SECONDARY_VAL_AX_ID = "-2113994441"

    def initialize(builder)
      super(nil, builder.chart_data)
      @builder = builder
    end

    def xml
      <<~XML
        #{ChartXmlWriter::Base::DECLARATION}
        <c:chartSpace #{ChartXmlWriter::Base::CHART_SPACE_NS}>
          <c:date1904 val="0"/>
          <c:chart>
            <c:autoTitleDeleted val="0"/>
            <c:plotArea>
              <c:layout/>
        #{indent(plot_elements, 6).chomp}
        #{indent(axes_xml, 6).chomp}
            </c:plotArea>
        #{indent(ChartXmlWriter::Base::LEGEND, 4).chomp}
            <c:plotVisOnly val="1"/>
            <c:dispBlanksAs val="gap"/>
            <c:showDLblsOverMax val="0"/>
          </c:chart>
        #{indent(ChartXmlWriter::Base::TEXT_PROPERTIES, 2).chomp}
        </c:chartSpace>
      XML
    end

    private

    def plot_elements = @builder.specs.map { |spec| plot_element(spec) }.join

    def plot_element(spec)
      family = ChartXmlWriter.family_of(spec.chart_type)
      value_ax = spec.secondary_axis ? SECONDARY_VAL_AX_ID : VAL_AX_ID
      cat_ax = spec.secondary_axis ? SECONDARY_CAT_AX_ID : CAT_AX_ID
      body = send(:"#{family}_plot_body", spec)

      <<~XML
        <#{plot_tag(family)}>
        #{indent(body, 2).chomp}
        #{indent(spec.series.map { |s| series_xml(s) }.join, 2).chomp}
        #{indent(plot_tail(family, spec), 2).chomp}
          <c:axId val="#{cat_ax}"/>
          <c:axId val="#{value_ax}"/>
        </#{plot_tag(family)}>
      XML
    end

    def plot_tag(family) = { bar: "c:barChart", line: "c:lineChart", area: "c:areaChart" }[family]

    def bar_plot_body(spec)
      direction = ChartXmlWriter::BarChartWriter::BAR_TYPES.include?(spec.chart_type.name) ? "bar" : "col"
      grouping = ChartXmlWriter::BarChartWriter::GROUPINGS.fetch(spec.chart_type.name)
      %(<c:barDir val="#{direction}"/>\n<c:grouping val="#{grouping}"/>\n)
    end

    def line_plot_body(spec)
      grouping = ChartXmlWriter::LineChartWriter::GROUPINGS.fetch(spec.chart_type.name)
      %(<c:grouping val="#{grouping}"/>\n<c:varyColors val="0"/>\n)
    end

    def area_plot_body(spec)
      grouping = ChartXmlWriter::AreaChartWriter::GROUPINGS.fetch(spec.chart_type.name)
      %(<c:grouping val="#{grouping}"/>\n<c:varyColors val="0"/>\n)
    end

    def plot_tail(family, spec)
      case family
      when :bar
        stacked = ChartXmlWriter::BarChartWriter::STACKED_TYPES.include?(spec.chart_type.name)
        stacked ? %(<c:overlap val="100"/>\n) : ""
      when :line
        %(<c:marker val="1"/>\n<c:smooth val="0"/>\n)
      else ""
      end
    end

    def axes_xml
      xml = +""
      xml << category_axis_xml(CAT_AX_ID, VAL_AX_ID, hidden: false)
      xml << value_axis_xml(VAL_AX_ID, CAT_AX_ID, position: "l", crosses: "autoZero")
      return xml unless @builder.secondary_axis?

      # The secondary pair: a value axis on the right, and the category axis
      # the schema demands to go with it, hidden because the primary pair
      # already draws the categories.
      xml << category_axis_xml(SECONDARY_CAT_AX_ID, SECONDARY_VAL_AX_ID, hidden: true)
      xml << value_axis_xml(SECONDARY_VAL_AX_ID, SECONDARY_CAT_AX_ID,
                            position: "r", crosses: "max")
      xml
    end

    def category_axis_xml(ax_id, cross_ax_id, hidden:)
      <<~XML
        <c:catAx>
          <c:axId val="#{ax_id}"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="#{hidden ? 1 : 0}"/>
          <c:axPos val="b"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="#{cross_ax_id}"/>
          <c:crosses val="autoZero"/>
          <c:auto val="1"/>
          <c:lblAlgn val="ctr"/>
          <c:lblOffset val="100"/>
          <c:noMultiLvlLbl val="0"/>
        </c:catAx>
      XML
    end

    def value_axis_xml(ax_id, cross_ax_id, position:, crosses:)
      <<~XML
        <c:valAx>
          <c:axId val="#{ax_id}"/>
          <c:scaling>
            <c:orientation val="minMax"/>
          </c:scaling>
          <c:delete val="0"/>
          <c:axPos val="#{position}"/>
          #{"<c:majorGridlines/>" if position == "l"}
          <c:numFmt formatCode="General" sourceLinked="1"/>
          <c:majorTickMark val="out"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="#{cross_ax_id}"/>
          <c:crosses val="#{crosses}"/>
        </c:valAx>
      XML
    end
  end
end
