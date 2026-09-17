# frozen_string_literal: true

require "pptx/enum/chart"
require "pptx/errors"

module Pptx
  # Builds the `c:chartSpace` XML for a new chart.
  #
  # Each chart family has its own skeleton, transcribed from the one
  # PowerPoint itself writes. The values are cached in this XML, which is what
  # PowerPoint renders from; the embedded workbook is only consulted when the
  # user edits the data.
  module ChartXmlWriter
    XL = Enum::XL_CHART_TYPE

    # Chart types this library can create, grouped by the plot element they
    # use. Anything outside this list raises rather than producing XML
    # PowerPoint would reject.
    FAMILIES = {
      bar: %i[BAR_CLUSTERED BAR_STACKED BAR_STACKED_100
              COLUMN_CLUSTERED COLUMN_STACKED COLUMN_STACKED_100],
      line: %i[LINE LINE_STACKED LINE_STACKED_100
               LINE_MARKERS LINE_MARKERS_STACKED LINE_MARKERS_STACKED_100],
      pie: %i[PIE],
      doughnut: %i[DOUGHNUT]
    }.freeze

    module_function

    # @param chart_type [Pptx::Enum::XL_CHART_TYPE]
    # @param chart_data [Pptx::ChartData]
    # @return [String] the chart part XML
    def write(chart_type, chart_data)
      member = XL.fetch(chart_type)
      case family_of(member)
      when :bar then BarChartWriter.new(member, chart_data).xml
      when :line then LineChartWriter.new(member, chart_data).xml
      when :pie then PieChartWriter.new(member, chart_data).xml
      when :doughnut then DoughnutChartWriter.new(member, chart_data).xml
      end
    end

    def family_of(member)
      FAMILIES.each { |family, names| return family if names.include?(member.name) }
      raise Error,
            "creating a #{member.name} chart is not supported yet; " \
            "supported types: #{FAMILIES.values.flatten.sort.join(', ')}"
    end

    def supported?(chart_type)
      member = XL[chart_type]
      !member.nil? && FAMILIES.values.any? { |names| names.include?(member.name) }
    end

    # Behaviour shared by the per-family writers.
    class Base
      DECLARATION = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
      CHART_SPACE_NS =
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' \
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' \
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

      TEXT_PROPERTIES = <<~XML
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr sz="1800"/>
            </a:pPr>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </c:txPr>
      XML

      LEGEND = <<~XML
        <c:legend>
          <c:legendPos val="r"/>
          <c:layout/>
          <c:overlay val="0"/>
        </c:legend>
      XML

      def initialize(chart_type, chart_data)
        @chart_type = chart_type
        @chart_data = chart_data
      end

      private

      attr_reader :chart_type, :chart_data

      # Ruby's squiggly heredoc only indents the first line of interpolated
      # content, so every fragment below is produced at zero indentation and
      # placed by its caller.
      def indent(text, spaces)
        pad = " " * spaces
        text.lines.map { |line| line.strip.empty? ? line : pad + line }.join
      end

      # One `c:ser` element: index, name, categories and values.
      def series_xml(series, extra_before_cat: "", extra_after_val: "")
        parts = [
          %(  <c:idx val="#{series.index}"/>\n),
          %(  <c:order val="#{series.index}"/>\n),
          indent(tx_xml(series), 2),
          extra_before_cat.empty? ? "" : indent(extra_before_cat, 2),
          indent(cat_xml(series), 2),
          indent(val_xml(series), 2),
          extra_after_val.empty? ? "" : indent(extra_after_val, 2)
        ]
        "<c:ser>\n#{parts.join}</c:ser>\n"
      end

      def tx_xml(series)
        <<~XML
          <c:tx>
            <c:strRef>
              <c:f>#{series.name_ref}</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
          #{indent(point_xml(0, escape(series.name)), 6).chomp}
              </c:strCache>
            </c:strRef>
          </c:tx>
        XML
      end

      def cat_xml(series)
        categories = chart_data.categories
        points = categories.each_with_index.map do |category, index|
          point_xml(index, format_value(category))
        end.join

        if chart_data.numeric_categories?
          <<~XML
            <c:cat>
              <c:numRef>
                <c:f>#{series.categories_ref}</c:f>
                <c:numCache>
                  <c:formatCode>#{chart_data.number_format}</c:formatCode>
                  <c:ptCount val="#{categories.size}"/>
            #{indent(points, 6).chomp}
                </c:numCache>
              </c:numRef>
            </c:cat>
          XML
        else
          <<~XML
            <c:cat>
              <c:strRef>
                <c:f>#{series.categories_ref}</c:f>
                <c:strCache>
                  <c:ptCount val="#{categories.size}"/>
            #{indent(points, 6).chomp}
                </c:strCache>
              </c:strRef>
            </c:cat>
          XML
        end
      end

      def val_xml(series)
        # A nil value means no data for that category: the point is omitted
        # while ptCount still counts it, which is how a gap is expressed.
        points = series.values.each_with_index.filter_map do |value, index|
          point_xml(index, format_value(value)) unless value.nil?
        end.join

        <<~XML
          <c:val>
            <c:numRef>
              <c:f>#{series.values_ref}</c:f>
              <c:numCache>
                <c:formatCode>#{series.number_format}</c:formatCode>
                <c:ptCount val="#{series.values.size}"/>
          #{indent(points, 6).chomp}
              </c:numCache>
            </c:numRef>
          </c:val>
        XML
      end

      def point_xml(index, value)
        %(<c:pt idx="#{index}">\n  <c:v>#{value}</c:v>\n</c:pt>\n)
      end

      # Floats keep their trailing ".0" here, unlike in the workbook, because
      # that is what the reference implementation writes.
      def format_value(value) = value.is_a?(String) ? escape(value) : value.to_s

      def escape(text)
        text.to_s.gsub("&", "&amp;").gsub("<", "&lt;").gsub(">", "&gt;")
      end

      def all_series_xml(&block)
        chart_data.series.map(&block).join
      end
    end

    # `c:barChart`: clustered, stacked and 100% stacked bars and columns.
    class BarChartWriter < Base
      CAT_AX_ID = "-2068027336"
      VAL_AX_ID = "-2113994440"

      BAR_TYPES = %i[BAR_CLUSTERED BAR_STACKED BAR_STACKED_100].freeze

      GROUPINGS = {
        BAR_CLUSTERED: "clustered", COLUMN_CLUSTERED: "clustered",
        BAR_STACKED: "stacked", COLUMN_STACKED: "stacked",
        BAR_STACKED_100: "percentStacked", COLUMN_STACKED_100: "percentStacked"
      }.freeze

      STACKED_TYPES = %i[BAR_STACKED BAR_STACKED_100 COLUMN_STACKED COLUMN_STACKED_100].freeze

      def xml
        <<~XML
          #{Base::DECLARATION}
          <c:chartSpace #{Base::CHART_SPACE_NS}>
            <c:date1904 val="0"/>
            <c:chart>
              <c:autoTitleDeleted val="0"/>
              <c:plotArea>
                <c:barChart>
                  <c:barDir val="#{bar_dir}"/>
                  <c:grouping val="#{grouping}"/>
          #{indent(series_blocks, 8).chomp}
          #{overlap}        <c:axId val="#{CAT_AX_ID}"/>
                  <c:axId val="#{VAL_AX_ID}"/>
                </c:barChart>
          #{indent(cat_ax_xml, 6).chomp}
                <c:valAx>
                  <c:axId val="#{VAL_AX_ID}"/>
                  <c:scaling/>
                  <c:delete val="0"/>
                  <c:axPos val="#{val_ax_pos}"/>
                  <c:majorGridlines/>
                  <c:majorTickMark val="out"/>
                  <c:minorTickMark val="none"/>
                  <c:tickLblPos val="nextTo"/>
                  <c:crossAx val="#{CAT_AX_ID}"/>
                  <c:crosses val="autoZero"/>
                </c:valAx>
              </c:plotArea>
              <c:dispBlanksAs val="gap"/>
            </c:chart>
          #{indent(Base::TEXT_PROPERTIES, 2).chomp}
          </c:chartSpace>
        XML
      end

      private

      def bar? = BAR_TYPES.include?(chart_type.name)

      def bar_dir = bar? ? "bar" : "col"

      def grouping = GROUPINGS.fetch(chart_type.name)

      # A bar chart's category axis runs down the left; a column chart's runs
      # along the bottom, and the value axis takes the other position.
      def cat_ax_pos = bar? ? "l" : "b"

      def val_ax_pos = bar? ? "b" : "l"

      def overlap
        STACKED_TYPES.include?(chart_type.name) ? %(        <c:overlap val="100"/>\n) : ""
      end

      def indent(text, spaces) = super

      def series_blocks = all_series_xml { |series| series_xml(series) }

      def cat_ax_xml
        <<~XML
          <c:catAx>
            <c:axId val="#{CAT_AX_ID}"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="#{cat_ax_pos}"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="#{VAL_AX_ID}"/>
            <c:crosses val="autoZero"/>
            <c:auto val="1"/>
            <c:lblAlgn val="ctr"/>
            <c:lblOffset val="100"/>
            <c:noMultiLvlLbl val="0"/>
          </c:catAx>
        XML
      end
    end

    # `c:lineChart`, with or without markers.
    class LineChartWriter < Base
      CAT_AX_ID = "2118791784"
      VAL_AX_ID = "2140495176"

      MARKER_TYPES = %i[LINE_MARKERS LINE_MARKERS_STACKED LINE_MARKERS_STACKED_100].freeze

      GROUPINGS = {
        LINE: "standard", LINE_MARKERS: "standard",
        LINE_STACKED: "stacked", LINE_MARKERS_STACKED: "stacked",
        LINE_STACKED_100: "percentStacked", LINE_MARKERS_STACKED_100: "percentStacked"
      }.freeze

      def xml
        <<~XML
          #{Base::DECLARATION}
          <c:chartSpace #{Base::CHART_SPACE_NS}>
            <c:date1904 val="0"/>
            <c:chart>
              <c:autoTitleDeleted val="0"/>
              <c:plotArea>
                <c:lineChart>
                  <c:grouping val="#{grouping}"/>
                  <c:varyColors val="0"/>
          #{indent(series_blocks, 8).chomp}
                  <c:marker val="1"/>
                  <c:smooth val="0"/>
                  <c:axId val="#{CAT_AX_ID}"/>
                  <c:axId val="#{VAL_AX_ID}"/>
                </c:lineChart>
                <c:catAx>
                  <c:axId val="#{CAT_AX_ID}"/>
                  <c:scaling>
                    <c:orientation val="minMax"/>
                  </c:scaling>
                  <c:delete val="0"/>
                  <c:axPos val="b"/>
                  <c:majorTickMark val="out"/>
                  <c:minorTickMark val="none"/>
                  <c:tickLblPos val="nextTo"/>
                  <c:crossAx val="#{VAL_AX_ID}"/>
                  <c:crosses val="autoZero"/>
                  <c:auto val="1"/>
                  <c:lblAlgn val="ctr"/>
                  <c:lblOffset val="100"/>
                  <c:noMultiLvlLbl val="0"/>
                </c:catAx>
                <c:valAx>
                  <c:axId val="#{VAL_AX_ID}"/>
                  <c:scaling/>
                  <c:delete val="0"/>
                  <c:axPos val="l"/>
                  <c:majorGridlines/>
                  <c:majorTickMark val="out"/>
                  <c:minorTickMark val="none"/>
                  <c:tickLblPos val="nextTo"/>
                  <c:crossAx val="#{CAT_AX_ID}"/>
                  <c:crosses val="autoZero"/>
                </c:valAx>
              </c:plotArea>
          #{indent(Base::LEGEND, 4).chomp}
              <c:plotVisOnly val="1"/>
              <c:dispBlanksAs val="gap"/>
              <c:showDLblsOverMax val="0"/>
            </c:chart>
          #{indent(Base::TEXT_PROPERTIES, 2).chomp}
          </c:chartSpace>
        XML
      end

      private

      def grouping = GROUPINGS.fetch(chart_type.name)

      def markers? = MARKER_TYPES.include?(chart_type.name)

      # A line chart without markers says so per series, by asking for the
      # "none" marker symbol.
      def marker_none_xml
        return "" if markers?

        <<~XML
          <c:marker>
            <c:symbol val="none"/>
          </c:marker>
        XML
      end

      def smooth_xml = %(<c:smooth val="0"/>\n)

      def series_blocks
        all_series_xml do |series|
          series_xml(series, extra_before_cat: marker_none_xml, extra_after_val: smooth_xml)
        end
      end
    end

    # `c:pieChart`, a single series shown as slices.
    class PieChartWriter < Base
      def xml
        <<~XML
          #{Base::DECLARATION}
          <c:chartSpace #{Base::CHART_SPACE_NS}>
            <c:chart>
              <c:autoTitleDeleted val="0"/>
              <c:plotArea>
                <c:pieChart>
                  <c:varyColors val="1"/>
          #{indent(series_blocks, 8).chomp}
                </c:pieChart>
              </c:plotArea>
              <c:dispBlanksAs val="gap"/>
            </c:chart>
          #{indent(Base::TEXT_PROPERTIES, 2).chomp}
          </c:chartSpace>
        XML
      end

      private

      # A pie plots one series only; any others live in the workbook but are
      # not drawn.
      def series_blocks = series_xml(chart_data.series.first)
    end

    # `c:doughnutChart`, a pie with a hole.
    class DoughnutChartWriter < Base
      # A doughnut's text properties leave the language off the trailing run
      # properties, where every other family sets it.
      TEXT_PROPERTIES = Base::TEXT_PROPERTIES.sub(
        %(<a:endParaRPr lang="en-US"/>), "<a:endParaRPr/>"
      )

      def xml
        <<~XML
          #{Base::DECLARATION}
          <c:chartSpace #{Base::CHART_SPACE_NS}>
            <c:date1904 val="0"/>
            <c:roundedCorners val="0"/>
            <c:chart>
              <c:autoTitleDeleted val="0"/>
              <c:plotArea>
                <c:layout/>
                <c:doughnutChart>
                  <c:varyColors val="1"/>
          #{indent(series_blocks, 8).chomp}
                  <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                    <c:showLeaderLines val="1"/>
                  </c:dLbls>
                  <c:firstSliceAng val="0"/>
                  <c:holeSize val="50"/>
                </c:doughnutChart>
              </c:plotArea>
          #{indent(Base::LEGEND, 4).chomp}
              <c:plotVisOnly val="1"/>
              <c:dispBlanksAs val="gap"/>
              <c:showDLblsOverMax val="0"/>
            </c:chart>
          #{indent(TEXT_PROPERTIES, 2).chomp}
          </c:chartSpace>
        XML
      end

      private

      # Unlike a pie, a doughnut draws every series as a concentric ring.
      def series_blocks = all_series_xml { |series| series_xml(series) }
    end
  end
end
