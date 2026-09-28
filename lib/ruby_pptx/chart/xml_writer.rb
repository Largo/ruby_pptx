# frozen_string_literal: true

require "ruby_pptx/enum/chart"
require "ruby_pptx/errors"

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
      xy: %i[XY_SCATTER XY_SCATTER_LINES XY_SCATTER_LINES_NO_MARKERS
             XY_SCATTER_SMOOTH XY_SCATTER_SMOOTH_NO_MARKERS],
      bubble: %i[BUBBLE BUBBLE_THREE_D_EFFECT],
      pie: %i[PIE],
      doughnut: %i[DOUGHNUT],
      area: %i[AREA AREA_STACKED AREA_STACKED_100],
      radar: %i[RADAR RADAR_MARKERS RADAR_FILLED]
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
      when :area then AreaChartWriter.new(member, chart_data).xml
      when :radar then RadarChartWriter.new(member, chart_data).xml
      when :xy then XyChartWriter.new(member, chart_data).xml
      when :bubble then BubbleChartWriter.new(member, chart_data).xml
      end
    end

    def family_of(member)
      FAMILIES.each { |family, names| return family if names.include?(member.name) }
      raise Error,
            "creating a #{member.name} chart is not supported yet; " \
            "supported types: #{FAMILIES.values.flatten.sort.join(", ")}"
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

      # Doughnut and area charts leave the language off the trailing run
      # properties where every other family sets it. There is no reason for
      # the difference beyond what the reference implementation writes, and a
      # byte comparison is the only thing that notices.
      TEXT_PROPERTIES_WITHOUT_LANG = TEXT_PROPERTIES.sub(
        %(<a:endParaRPr lang="en-US"/>), "<a:endParaRPr/>"
      )

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

      # Protected rather than private: ComboChartWriter builds its plots from
      # the same fragments, and is a sibling rather than a subclass of the
      # per-family writers.
      protected

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

      # The categories: a numeric cache for numbers and dates, a string cache
      # for a plain list, and a multi-level cache for grouped categories.
      def cat_xml(series)
        categories = chart_data.categories
        if categories.numeric?
          numeric_cat_xml(series, categories)
        elsif categories.depth == 1
          string_cat_xml(series, categories)
        else
          multi_level_cat_xml(series, categories)
        end
      end

      def val_xml(series)
        num_ref_xml("c:val", series.values_ref, series.values, series.number_format)
      end

      # A `c:val`, `c:xVal`, `c:yVal` or `c:bubbleSize` block: a worksheet
      # reference plus the cached values.
      #
      # A nil value means no data at that position: the point is omitted while
      # ptCount still counts it, which is how a gap is expressed.
      def num_ref_xml(tag, reference, values, number_format)
        points = values.each_with_index.filter_map do |value, index|
          point_xml(index, format_value(value)) unless value.nil?
        end.join

        <<~XML
          <#{tag}>
            <c:numRef>
              <c:f>#{reference}</c:f>
              <c:numCache>
                <c:formatCode>#{number_format}</c:formatCode>
                <c:ptCount val="#{values.size}"/>
          #{indent(points, 6).chomp}
              </c:numCache>
            </c:numRef>
          </#{tag}>
        XML
      end

      def numeric_cat_xml(series, categories)
        points = categories.each_with_index.map do |category, index|
          point_xml(index, category.numeric_str_val(date_1904: date_1904?))
        end.join
        <<~XML
          <c:cat>
            <c:numRef>
              <c:f>#{series.categories_ref}</c:f>
              <c:numCache>
                <c:formatCode>#{categories.number_format}</c:formatCode>
                <c:ptCount val="#{categories.leaf_count}"/>
          #{indent(points, 6).chomp}
              </c:numCache>
            </c:numRef>
          </c:cat>
        XML
      end

      def string_cat_xml(series, categories)
        points = categories.each_with_index.map do |category, index|
          point_xml(index, escape(category.label))
        end.join
        <<~XML
          <c:cat>
            <c:strRef>
              <c:f>#{series.categories_ref}</c:f>
              <c:strCache>
                <c:ptCount val="#{categories.leaf_count}"/>
          #{indent(points, 6).chomp}
              </c:strCache>
            </c:strRef>
          </c:cat>
        XML
      end

      # One `c:lvl` per level, leaf level first. Each label sits at the leaf
      # index where its run begins, so a parent's points are sparse.
      def multi_level_cat_xml(series, categories)
        levels = categories.levels.map do |level|
          points = level.map { |index, label| point_xml(index, escape(label)) }.join
          "<c:lvl>\n#{indent(points, 2)}</c:lvl>\n"
        end.join
        <<~XML
          <c:cat>
            <c:multiLvlStrRef>
              <c:f>#{series.categories_ref}</c:f>
              <c:multiLvlStrCache>
                <c:ptCount val="#{categories.leaf_count}"/>
          #{indent(levels, 6).chomp}
              </c:multiLvlStrCache>
            </c:multiLvlStrRef>
          </c:cat>
        XML
      end

      # The axis written instead of a category axis when the categories are
      # dates. Bar, line and area charts share it, differing only in ids and
      # position.
      def date_ax_xml(ax_id, cross_ax_id, position)
        <<~XML
          <c:dateAx>
            <c:axId val="#{ax_id}"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="#{position}"/>
            <c:numFmt formatCode="#{chart_data.categories.number_format}" sourceLinked="1"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="#{cross_ax_id}"/>
            <c:crosses val="autoZero"/>
            <c:auto val="1"/>
            <c:lblOffset val="100"/>
            <c:baseTimeUnit val="days"/>
          </c:dateAx>
        XML
      end

      def dates?
        chart_data.categories.dates?
      end

      # A chart being written from scratch uses the 1900 date system;
      # rewriting an existing chart's data honours the chart's own setting.
      def date_1904?
        @date_1904 || false
      end

      def point_xml(index, value)
        %(<c:pt idx="#{index}">\n  <c:v>#{value}</c:v>\n</c:pt>\n)
      end

      # Floats keep their trailing ".0" here, unlike in the workbook, because
      # that is what the reference implementation writes.
      def format_value(value)
        value.is_a?(String) ? escape(value) : value.to_s
      end

      def escape(text)
        text.to_s.gsub("&", "&amp;").gsub("<", "&lt;").gsub(">", "&gt;")
      end

      def all_series_xml(&)
        chart_data.series.map(&).join
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

      def bar?
        BAR_TYPES.include?(chart_type.name)
      end

      def bar_dir
        bar? ? "bar" : "col"
      end

      def grouping
        GROUPINGS.fetch(chart_type.name)
      end

      # A bar chart's category axis runs down the left; a column chart's runs
      # along the bottom, and the value axis takes the other position.
      def cat_ax_pos
        bar? ? "l" : "b"
      end

      def val_ax_pos
        bar? ? "b" : "l"
      end

      def overlap
        STACKED_TYPES.include?(chart_type.name) ? %(        <c:overlap val="100"/>\n) : ""
      end

      def series_blocks
        all_series_xml { |series| series_xml(series) }
      end

      def cat_ax_xml
        return date_ax_xml(CAT_AX_ID, VAL_AX_ID, cat_ax_pos) if dates?

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
          #{indent(cat_ax_xml, 6).chomp}
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

      def grouping
        GROUPINGS.fetch(chart_type.name)
      end

      def markers?
        MARKER_TYPES.include?(chart_type.name)
      end

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

      def smooth_xml
        %(<c:smooth val="0"/>\n)
      end

      def series_blocks
        all_series_xml do |series|
          series_xml(series, extra_before_cat: marker_none_xml, extra_after_val: smooth_xml)
        end
      end

      def cat_ax_xml
        return date_ax_xml(CAT_AX_ID, VAL_AX_ID, "b") if dates?

        <<~XML
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
        XML
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
      def series_blocks
        series_xml(chart_data.series.first)
      end
    end

    # Shared by the two families whose points carry their own x value rather
    # than sitting against a shared category.
    class XyBase < Base
      private

      # A scatter or bubble series names its own x and y ranges instead of
      # sharing a category axis.
      def xy_series_xml(series, extra_after_tx: "", extra_after_values: "")
        parts = [
          %(  <c:idx val="#{series.index}"/>\n),
          %(  <c:order val="#{series.index}"/>\n),
          indent(tx_xml(series), 2),
          extra_after_tx.empty? ? "" : indent(extra_after_tx, 2),
          indent(num_ref_xml("c:xVal", series.x_values_ref, series.x_values,
                             series.number_format), 2),
          indent(num_ref_xml("c:yVal", series.y_values_ref, series.y_values,
                             series.number_format), 2),
          extra_after_values.empty? ? "" : indent(extra_after_values, 2)
        ]
        "<c:ser>\n#{parts.join}</c:ser>\n"
      end

      def value_axes_xml(x_ax_id, y_ax_id)
        <<~XML
          <c:valAx>
            <c:axId val="#{x_ax_id}"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="b"/>
            <c:numFmt formatCode="General" sourceLinked="1"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="#{y_ax_id}"/>
            <c:crosses val="autoZero"/>
            <c:crossBetween val="midCat"/>
          </c:valAx>
          <c:valAx>
            <c:axId val="#{y_ax_id}"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="l"/>
            <c:majorGridlines/>
            <c:numFmt formatCode="General" sourceLinked="1"/>
            <c:majorTickMark val="out"/>
            <c:minorTickMark val="none"/>
            <c:tickLblPos val="nextTo"/>
            <c:crossAx val="#{x_ax_id}"/>
            <c:crosses val="autoZero"/>
            <c:crossBetween val="midCat"/>
          </c:valAx>
        XML
      end
    end

    # `c:scatterChart`, points at (x, y) with optional lines.
    class XyChartWriter < XyBase
      X_AX_ID = "-2128940872"
      Y_AX_ID = "-2129643912"

      SMOOTH_TYPES = %i[XY_SCATTER_SMOOTH XY_SCATTER_SMOOTH_NO_MARKERS].freeze
      NO_MARKER_TYPES = %i[XY_SCATTER_LINES_NO_MARKERS XY_SCATTER_SMOOTH_NO_MARKERS].freeze

      def xml
        <<~XML
          #{Base::DECLARATION}
          <c:chartSpace #{Base::CHART_SPACE_NS}>
            <c:chart>
              <c:plotArea>
                <c:scatterChart>
                  <c:scatterStyle val="#{scatter_style}"/>
                  <c:varyColors val="0"/>
          #{indent(series_blocks, 8).chomp}
                  <c:axId val="#{X_AX_ID}"/>
                  <c:axId val="#{Y_AX_ID}"/>
                </c:scatterChart>
          #{indent(value_axes_xml(X_AX_ID, Y_AX_ID), 6).chomp}
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

      def scatter_style
        SMOOTH_TYPES.include?(chart_type.name) ? "smoothMarker" : "lineMarker"
      end

      # Plain XY_SCATTER draws markers only, which is said by giving the series
      # an invisible line rather than by any scatterStyle value.
      def line_suppression_xml
        return "" unless chart_type.name == :XY_SCATTER

        <<~XML
          <c:spPr>
            <a:ln w="47625">
              <a:noFill/>
            </a:ln>
          </c:spPr>
        XML
      end

      def marker_none_xml
        return "" unless NO_MARKER_TYPES.include?(chart_type.name)

        <<~XML
          <c:marker>
            <c:symbol val="none"/>
          </c:marker>
        XML
      end

      def series_blocks
        all_series_xml do |series|
          xy_series_xml(series,
                        extra_after_tx: line_suppression_xml + marker_none_xml,
                        extra_after_values: %(<c:smooth val="0"/>\n))
        end
      end
    end

    # `c:bubbleChart`, points at (x, y) sized by a third value.
    class BubbleChartWriter < XyBase
      X_AX_ID = "-2115720072"
      Y_AX_ID = "-2115723560"

      def xml
        <<~XML
          #{Base::DECLARATION}
          <c:chartSpace #{Base::CHART_SPACE_NS}>
            <c:chart>
              <c:autoTitleDeleted val="0"/>
              <c:plotArea>
                <c:layout/>
                <c:bubbleChart>
                  <c:varyColors val="0"/>
          #{indent(series_blocks, 8).chomp}
                  <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                  </c:dLbls>
                  <c:bubbleScale val="100"/>
                  <c:showNegBubbles val="0"/>
                  <c:axId val="#{X_AX_ID}"/>
                  <c:axId val="#{Y_AX_ID}"/>
                </c:bubbleChart>
          #{indent(value_axes_xml(X_AX_ID, Y_AX_ID), 6).chomp}
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

      def three_d?
        chart_type.name == :BUBBLE_THREE_D_EFFECT
      end

      def series_blocks
        all_series_xml do |series|
          sizes = indent(num_ref_xml("c:bubbleSize", series.bubble_sizes_ref,
                                     series.bubble_sizes, series.number_format), 0)
          xy_series_xml(series,
                        extra_after_tx: %(<c:invertIfNegative val="0"/>\n),
                        extra_after_values: sizes + %(<c:bubble3D val="#{three_d? ? 1 : 0}"/>\n))
        end
      end
    end

    # `c:areaChart`, filled bands under each series.
    class AreaChartWriter < Base
      CAT_AX_ID = "-2101159928"
      VAL_AX_ID = "-2100718248"

      GROUPINGS = {
        AREA: "standard", AREA_STACKED: "stacked", AREA_STACKED_100: "percentStacked"
      }.freeze

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
                <c:areaChart>
                  <c:grouping val="#{grouping}"/>
                  <c:varyColors val="0"/>
          #{indent(series_blocks, 8).chomp}
                  <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                  </c:dLbls>
                  <c:axId val="#{CAT_AX_ID}"/>
                  <c:axId val="#{VAL_AX_ID}"/>
                </c:areaChart>
          #{indent(cat_ax_xml, 6).chomp}
                <c:valAx>
                  <c:axId val="#{VAL_AX_ID}"/>
                  <c:scaling>
                    <c:orientation val="minMax"/>
                  </c:scaling>
                  <c:delete val="0"/>
                  <c:axPos val="l"/>
                  <c:majorGridlines/>
                  <c:numFmt formatCode="General" sourceLinked="1"/>
                  <c:majorTickMark val="out"/>
                  <c:minorTickMark val="none"/>
                  <c:tickLblPos val="nextTo"/>
                  <c:crossAx val="#{CAT_AX_ID}"/>
                  <c:crosses val="autoZero"/>
                  <c:crossBetween val="midCat"/>
                </c:valAx>
              </c:plotArea>
          #{indent(Base::LEGEND, 4).chomp}
              <c:plotVisOnly val="1"/>
              <c:dispBlanksAs val="zero"/>
              <c:showDLblsOverMax val="0"/>
            </c:chart>
          #{indent(Base::TEXT_PROPERTIES_WITHOUT_LANG, 2).chomp}
          </c:chartSpace>
        XML
      end

      private

      def grouping
        GROUPINGS.fetch(chart_type.name)
      end

      def series_blocks
        all_series_xml { |series| series_xml(series) }
      end

      def cat_ax_xml
        return date_ax_xml(CAT_AX_ID, VAL_AX_ID, "b") if dates?

        <<~XML
          <c:catAx>
            <c:axId val="#{CAT_AX_ID}"/>
            <c:scaling>
              <c:orientation val="minMax"/>
            </c:scaling>
            <c:delete val="0"/>
            <c:axPos val="b"/>
            <c:numFmt formatCode="General" sourceLinked="1"/>
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

    # `c:radarChart`, values plotted on spokes around a centre.
    class RadarChartWriter < Base
      CAT_AX_ID = "2073612648"
      VAL_AX_ID = "-2112772216"

      # A radar chart carries a style hint PowerPoint 2007 did not understand,
      # so it is written inside an AlternateContent block with a plain
      # fallback.
      ALTERNATE_CONTENT = <<~XML
        <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14">
            <c14:style val="118"/>
          </mc:Choice>
          <mc:Fallback>
            <c:style val="18"/>
          </mc:Fallback>
        </mc:AlternateContent>
      XML

      def xml
        <<~XML
          #{Base::DECLARATION}
          <c:chartSpace #{Base::CHART_SPACE_NS}>
            <c:date1904 val="0"/>
            <c:roundedCorners val="0"/>
          #{indent(ALTERNATE_CONTENT, 2).chomp}
            <c:chart>
              <c:autoTitleDeleted val="0"/>
              <c:plotArea>
                <c:layout/>
                <c:radarChart>
                  <c:radarStyle val="#{radar_style}"/>
                  <c:varyColors val="0"/>
          #{indent(series_blocks, 8).chomp}
                  <c:axId val="#{CAT_AX_ID}"/>
                  <c:axId val="#{VAL_AX_ID}"/>
                </c:radarChart>
                <c:catAx>
                  <c:axId val="#{CAT_AX_ID}"/>
                  <c:scaling>
                    <c:orientation val="minMax"/>
                  </c:scaling>
                  <c:delete val="0"/>
                  <c:axPos val="b"/>
                  <c:majorGridlines/>
                  <c:numFmt formatCode="m/d/yy" sourceLinked="1"/>
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
                  <c:scaling>
                    <c:orientation val="minMax"/>
                  </c:scaling>
                  <c:delete val="0"/>
                  <c:axPos val="l"/>
                  <c:majorGridlines/>
                  <c:numFmt formatCode="General" sourceLinked="1"/>
                  <c:majorTickMark val="cross"/>
                  <c:minorTickMark val="none"/>
                  <c:tickLblPos val="nextTo"/>
                  <c:crossAx val="#{CAT_AX_ID}"/>
                  <c:crosses val="autoZero"/>
                  <c:crossBetween val="between"/>
                </c:valAx>
              </c:plotArea>
              <c:plotVisOnly val="1"/>
              <c:dispBlanksAs val="gap"/>
              <c:showDLblsOverMax val="0"/>
            </c:chart>
          #{indent(Base::TEXT_PROPERTIES, 2).chomp}
          </c:chartSpace>
        XML
      end

      private

      def radar_style
        chart_type.name == :RADAR_FILLED ? "filled" : "marker"
      end

      # A plain radar draws lines without markers, which is said per series by
      # asking for the "none" symbol. The markers and filled variants do not.
      def marker_none_xml
        return "" unless chart_type.name == :RADAR

        <<~XML
          <c:marker>
            <c:symbol val="none"/>
          </c:marker>
        XML
      end

      def series_blocks
        all_series_xml do |series|
          series_xml(series, extra_before_cat: marker_none_xml,
                             extra_after_val: %(<c:smooth val="0"/>\n))
        end
      end
    end

    # `c:doughnutChart`, a pie with a hole.
    class DoughnutChartWriter < Base
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
          #{indent(Base::TEXT_PROPERTIES_WITHOUT_LANG, 2).chomp}
          </c:chartSpace>
        XML
      end

      private

      # Unlike a pie, a doughnut draws every series as a concentric ring.
      def series_blocks
        all_series_xml { |series| series_xml(series) }
      end
    end
  end
end
