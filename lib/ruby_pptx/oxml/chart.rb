# frozen_string_literal: true

require "ruby_pptx/oxml/element"
require "ruby_pptx/oxml/content_model"
require "ruby_pptx/oxml/simple_types"
require "ruby_pptx/oxml/text"
require "ruby_pptx/enum/chart"

module Pptx
  module Oxml
    # Chart elements, after python-pptx's `oxml/chart` package.
    #
    # The class boundaries, tag sequences and the split between the two kinds
    # of boolean element follow python-pptx's registrations exactly: what a
    # chart property writes depends on which class its element belongs to,
    # and byte parity with python-pptx is the point.

    # -------------------------------------------------------------------------
    # Value elements -- one `val` attribute, typed by what it holds
    # -------------------------------------------------------------------------

    # A boolean `val` whose schema default is true, so writing true removes
    # the attribute: `<c:smooth/>` means smooth. Absent reads as true.
    class CT_Boolean < Element
      tag "c:delete", "c:varyColors", "c:smooth", "c:bubble3D", "c:date1904"
      optional_attr "val", type: SimpleTypes::XsdBoolean, default: true
    end

    # A boolean `val` always written out, `val="1"` or `val="0"`, even though
    # the schema default is the same true. PowerPoint writes these explicitly
    # and python-pptx does too.
    class CT_BooleanExplicit < Element
      tag "c:overlay", "c:showVal", "c:showCatName", "c:showSerName", "c:showPercent",
          "c:showLegendKey", "c:showBubbleSize", "c:autoTitleDeleted", "c:invertIfNegative",
          "c:plotVisOnly", "c:showDLblsOverMax", "c:showLeaderLines"

      def val
        raw = get("val")
        raw.nil? || SimpleTypes::XsdBoolean.from_xml(raw)
      end

      def val=(value)
        set("val", value ? "1" : "0")
      end
    end

    # `c:autoUpdate`, whether the chart refreshes itself from the workbook.
    class CT_AutoUpdate < Element
      tag "c:autoUpdate"
      optional_attr "val", type: SimpleTypes::XsdBoolean
    end

    # A double-valued `val`: axis scale limits, a crossing point, a manual
    # layout coordinate.
    class CT_Double < Element
      tag "c:max", "c:min", "c:crossesAt", "c:x"
      required_attr "val", type: SimpleTypes::XsdDouble
    end

    # `c:majorUnit` and `c:minorUnit`, which must be positive.
    class CT_AxisUnit < Element
      tag "c:majorUnit", "c:minorUnit"
      required_attr "val", type: SimpleTypes::ST_AxisUnit
    end

    # An unsigned integer `val`: series index and order, axis ids.
    class CT_UnsignedInt < Element
      tag "c:idx", "c:order", "c:crossAx", "c:axId", "c:ptCount"
      required_attr "val", type: SimpleTypes::XsdUnsignedInt
    end

    # String-valued elements this library writes from templates but has no
    # typed accessor for.
    class CT_ChartString < Element
      tag "c:axPos", "c:radarStyle", "c:scatterStyle", "c:dispBlanksAs"
      optional_attr "val", type: SimpleTypes::XsdString
    end

    class CT_LegendPos < Element
      tag "c:legendPos"
      optional_attr "val", type: Enum::XL_LEGEND_POSITION, default: Enum::XL_LEGEND_POSITION::RIGHT
    end

    class CT_DLblPos < Element
      tag "c:dLblPos"
      required_attr "val", type: Enum::XL_DATA_LABEL_POSITION
    end

    class CT_TickLblPos < Element
      tag "c:tickLblPos"
      optional_attr "val", type: Enum::XL_TICK_LABEL_POSITION
    end

    class CT_TickMark < Element
      tag "c:majorTickMark", "c:minorTickMark"
      optional_attr "val", type: Enum::XL_TICK_MARK, default: Enum::XL_TICK_MARK::CROSS
    end

    class CT_Crosses < Element
      tag "c:crosses"
      required_attr "val", type: Enum::XL_AXIS_CROSSES
    end

    class CT_Orientation < Element
      tag "c:orientation"
      optional_attr "val", type: SimpleTypes::ST_Orientation,
                           default: SimpleTypes::ST_Orientation::MIN_MAX
    end

    class CT_LblOffset < Element
      tag "c:lblOffset"
      optional_attr "val", type: SimpleTypes::ST_LblOffset, default: 100
    end

    class CT_Grouping < Element
      tag "c:grouping"
      optional_attr "val", type: SimpleTypes::ST_Grouping
    end

    class CT_BarDir < Element
      tag "c:barDir"
      optional_attr "val", type: SimpleTypes::ST_BarDir, default: SimpleTypes::ST_BarDir::COL
    end

    # `c:gapWidth`: space between bar groups, as a percentage of bar width.
    class CT_GapAmount < Element
      tag "c:gapWidth"
      optional_attr "val", type: SimpleTypes::ST_GapAmount, default: 150
    end

    class CT_Overlap < Element
      tag "c:overlap"
      required_attr "val", type: SimpleTypes::ST_Overlap
    end

    class CT_BubbleScale < Element
      tag "c:bubbleScale"
      optional_attr "val", type: SimpleTypes::ST_BubbleScale, default: 100
    end

    # `c:style`, the chart style number shown in PowerPoint's style gallery.
    class CT_Style < Element
      tag "c:style"
      required_attr "val", type: SimpleTypes::ST_Style
    end

    class CT_MarkerSize < Element
      tag "c:size"
      required_attr "val", type: SimpleTypes::ST_MarkerSize
    end

    class CT_MarkerStyle < Element
      tag "c:symbol"
      required_attr "val", type: Enum::XL_MARKER_STYLE
    end

    class CT_LayoutMode < Element
      tag "c:xMode"
      optional_attr "val", type: SimpleTypes::ST_LayoutMode,
                           default: SimpleTypes::ST_LayoutMode::FACTOR
    end

    # `c:numFmt`, a number format applied to an axis or to data labels.
    class CT_NumberFormat < Element
      tag "c:numFmt"
      optional_attr "formatCode", type: SimpleTypes::XsdString
      optional_attr "sourceLinked", type: SimpleTypes::XsdBoolean
    end

    # -------------------------------------------------------------------------
    # Shared structure
    # -------------------------------------------------------------------------

    # Anything carrying a `c:txPr`, which is where a chart element's font
    # lives. The text body is created on first use.
    module ChartTextProperties
      def defRPr
        get_or_add_txPr.defRPr
      end

      def new_txPr
        build_from_xml(CT_TextBody::TXPR_XML)
      end
    end

    # `c:layout`, holding an optional manual layout.
    class CT_Layout < Element
      tag "c:layout"
      zero_or_one "c:manualLayout", successors: %w[c:extLst]

      # The horizontal offset as a fraction of the chart width; 0.0 when the
      # position is automatic.
      def horz_offset
        manualLayout&.horz_offset || 0.0
      end

      # Setting 0.0 hands the position back to PowerPoint.
      def horz_offset=(offset)
        if offset.zero?
          remove_manualLayout
          return
        end

        get_or_add_manualLayout.horz_offset = offset
      end
    end

    class CT_ManualLayout < Element
      tag "c:manualLayout"
      TAG_SEQ = %w[c:layoutTarget c:xMode c:yMode c:wMode c:hMode c:x c:y c:w c:h
                   c:extLst].freeze
      zero_or_one "c:xMode", successors: TAG_SEQ[2..]
      zero_or_one "c:x", successors: TAG_SEQ[6..]

      # Only a factor-mode x is an offset; an edge-mode x is a position.
      def horz_offset
        return 0.0 if x.nil? || xMode.nil? || xMode.val != SimpleTypes::ST_LayoutMode::FACTOR

        x.val
      end

      def horz_offset=(offset)
        get_or_add_xMode.val = SimpleTypes::ST_LayoutMode::FACTOR
        get_or_add_x.val = offset
      end
    end

    # `c:tx`: a title's or data label's rich text, or a series name's cell
    # reference. The same element serves all three, so one class does.
    class CT_Tx < Element
      tag "c:tx"
      zero_or_one "c:strRef", successors: []
      zero_or_one "c:rich", successors: []

      RICH_XML = <<~XML.freeze
        <c:rich #{Ns.nsdecls("c", "a")}>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr>
              <a:defRPr/>
            </a:pPr>
          </a:p>
        </c:rich>
      XML

      def new_rich
        build_from_xml(RICH_XML)
      end
    end

    # `c:title`, the chart or axis title.
    class CT_Title < Element
      tag "c:title"
      TAG_SEQ = %w[c:tx c:layout c:overlay c:spPr c:txPr c:extLst].freeze

      zero_or_one "c:tx", successors: TAG_SEQ[1..]
      zero_or_one "c:spPr", successors: TAG_SEQ[4..]

      # A title with no text of its own yet, which is what PowerPoint and
      # python-pptx both write when a title is switched on.
      def self.new_title(context)
        context.build_from_xml(<<~XML)
          <c:title #{Ns.nsdecls("c")}>
            <c:layout/>
            <c:overlay val="0"/>
          </c:title>
        XML
      end

      # The rich-text body, or nil when the title has none.
      def rich
        xpath("./c:tx/c:rich").first
      end

      # A title takes rich text or a cell reference, never both.
      def get_or_add_rich
        tx = get_or_add_tx
        tx.remove_strRef
        tx.get_or_add_rich
      end
    end

    # `c:majorGridlines` and `c:minorGridlines`.
    class CT_ChartLines < Element
      tag "c:majorGridlines", "c:minorGridlines"
      zero_or_one "c:spPr", successors: []
    end

    # -------------------------------------------------------------------------
    # Axes
    # -------------------------------------------------------------------------

    # `c:scaling`, the axis range and direction.
    class CT_Scaling < Element
      tag "c:scaling"
      zero_or_one "c:logBase", successors: %w[c:orientation c:max c:min c:extLst]
      zero_or_one "c:orientation", successors: %w[c:max c:min c:extLst]
      zero_or_one "c:max", successors: %w[c:min c:extLst]
      zero_or_one "c:min", successors: %w[c:extLst]

      # nil means "auto", which is what an absent element says.
      def maximum
        max&.val
      end

      def maximum=(value)
        remove_max
        get_or_add_max.val = value unless value.nil?
      end

      def minimum
        min&.val
      end

      def minimum=(value)
        remove_min
        get_or_add_min.val = value unless value.nil?
      end
    end

    # What the three axis elements share. They agree up to `c:crossesAt` and
    # then diverge, which is why each declares its own tail below.
    module BaseAxisElement
      include ChartTextProperties

      def self.included(base)
        base.class_eval do
          seq = self::TAG_SEQ
          one_and_only_one "c:scaling"
          zero_or_one "c:delete", successors: seq[3..]
          zero_or_one "c:majorGridlines", successors: seq[5..]
          zero_or_one "c:minorGridlines", successors: seq[6..]
          zero_or_one "c:title", successors: seq[7..]
          zero_or_one "c:numFmt", successors: seq[8..]
          zero_or_one "c:majorTickMark", successors: seq[9..]
          zero_or_one "c:minorTickMark", successors: seq[10..]
          zero_or_one "c:tickLblPos", successors: seq[11..]
          zero_or_one "c:spPr", successors: seq[12..]
          zero_or_one "c:txPr", successors: seq[13..]
          zero_or_one "c:crossAx", successors: seq[14..]
          zero_or_one "c:crosses", successors: seq[15..]
          zero_or_one "c:crossesAt", successors: seq[16..]
        end
      end

      def orientation
        scaling.orientation&.val || SimpleTypes::ST_Orientation::MIN_MAX
      end

      # Only a reversed axis is written; the normal direction is the default.
      def orientation=(value)
        scaling.remove_orientation
        return unless value == SimpleTypes::ST_Orientation::MAX_MIN

        scaling.get_or_add_orientation.val = value
      end

      def new_title
        CT_Title.new_title(self)
      end
    end

    COMMON_AXIS_SEQ = %w[c:axId c:scaling c:delete c:axPos c:majorGridlines c:minorGridlines
                         c:title c:numFmt c:majorTickMark c:minorTickMark c:tickLblPos c:spPr
                         c:txPr c:crossAx c:crosses c:crossesAt].freeze

    class CT_CatAx < Element
      tag "c:catAx"
      TAG_SEQ = (COMMON_AXIS_SEQ + %w[c:auto c:lblAlgn c:lblOffset c:tickLblSkip
                                      c:tickMarkSkip c:noMultiLvlLbl c:extLst]).freeze
      include BaseAxisElement

      zero_or_one "c:lblOffset", successors: TAG_SEQ[19..]
    end

    class CT_DateAx < Element
      tag "c:dateAx"
      TAG_SEQ = (COMMON_AXIS_SEQ + %w[c:auto c:lblOffset c:baseTimeUnit c:majorUnit
                                      c:majorTimeUnit c:minorUnit c:minorTimeUnit
                                      c:extLst]).freeze
      include BaseAxisElement

      zero_or_one "c:lblOffset", successors: TAG_SEQ[18..]
    end

    class CT_ValAx < Element
      tag "c:valAx"
      TAG_SEQ = (COMMON_AXIS_SEQ + %w[c:crossBetween c:majorUnit c:minorUnit c:dispUnits
                                      c:extLst]).freeze
      include BaseAxisElement

      zero_or_one "c:majorUnit", successors: TAG_SEQ[18..]
      zero_or_one "c:minorUnit", successors: TAG_SEQ[19..]
    end

    # -------------------------------------------------------------------------
    # Legend
    # -------------------------------------------------------------------------

    class CT_Legend < Element
      include ChartTextProperties

      tag "c:legend"
      TAG_SEQ = %w[c:legendPos c:legendEntry c:layout c:overlay c:spPr c:txPr c:extLst].freeze

      zero_or_one "c:legendPos", successors: TAG_SEQ[1..]
      zero_or_one "c:layout", successors: TAG_SEQ[3..]
      zero_or_one "c:overlay", successors: TAG_SEQ[4..]
      zero_or_one "c:txPr", successors: TAG_SEQ[6..]

      def horz_offset
        layout&.horz_offset || 0.0
      end

      def horz_offset=(offset)
        get_or_add_layout.horz_offset = offset
      end
    end

    # -------------------------------------------------------------------------
    # Data labels
    # -------------------------------------------------------------------------

    # `c:dLbls`, data-label settings for a plot or a series.
    class CT_DataLabels < Element
      include ChartTextProperties

      tag "c:dLbls"
      TAG_SEQ = %w[c:dLbl c:numFmt c:spPr c:txPr c:dLblPos c:showLegendKey c:showVal
                   c:showCatName c:showSerName c:showPercent c:showBubbleSize
                   c:separator c:showLeaderLines c:leaderLines c:extLst].freeze

      zero_or_more "c:dLbl", successors: TAG_SEQ[1..], as: :dLbl
      zero_or_one "c:numFmt", successors: TAG_SEQ[2..]
      zero_or_one "c:txPr", successors: TAG_SEQ[4..]
      zero_or_one "c:dLblPos", successors: TAG_SEQ[5..]
      zero_or_one "c:showLegendKey", successors: TAG_SEQ[6..]
      zero_or_one "c:showVal", successors: TAG_SEQ[7..]
      zero_or_one "c:showCatName", successors: TAG_SEQ[8..]
      zero_or_one "c:showSerName", successors: TAG_SEQ[9..]
      zero_or_one "c:showPercent", successors: TAG_SEQ[10..]
      zero_or_one "c:showBubbleSize", successors: TAG_SEQ[11..]
      zero_or_one "c:showLeaderLines", successors: TAG_SEQ[13..]

      # What PowerPoint writes when labels are switched on: every kind of
      # label named and turned off, leader lines allowed.
      DEFAULT_XML = <<~XML.freeze
        <c:dLbls #{Ns.nsdecls("c")}>
          <c:showLegendKey val="0"/>
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
          <c:showLeaderLines val="1"/>
        </c:dLbls>
      XML

      def self.new_data_labels(context)
        context.build_from_xml(DEFAULT_XML)
      end

      # A flag added on its own starts off, as PowerPoint writes them.
      %w[showLegendKey showVal showCatName showSerName showPercent].each do |flag|
        define_method(:"new_#{flag}") { build_from_xml(%(<c:#{flag} #{Ns.nsdecls("c")} val="0"/>)) }
      end

      # The label for the point at +idx+, or nil if it has none of its own.
      def dLbl_for_point(idx)
        dLbl_list.find { |label| label.idx.val == idx }
      end

      # The label for the point at +idx+, created in index order if absent.
      def get_or_add_dLbl_for_point(idx)
        dLbl_for_point(idx) || insert_dLbl_in_sequence(idx)
      end

      private

      def insert_dLbl_in_sequence(idx)
        label = CT_DLbl.new_dLbl(self)
        label.idx.val = idx
        following = dLbl_list.find { |existing| existing.idx.val > idx }
        if following
          following.node.add_previous_sibling(label.node)
        elsif (last = dLbl_list.last)
          last.node.add_next_sibling(label.node)
        else
          node.prepend_child(label.node)
        end
        label
      end
    end

    # `c:dLbl`, the label of one data point.
    class CT_DLbl < Element
      include ChartTextProperties

      tag "c:dLbl"
      TAG_SEQ = %w[c:idx c:layout c:tx c:numFmt c:spPr c:txPr c:dLblPos c:showLegendKey
                   c:showVal c:showCatName c:showSerName c:showPercent c:showBubbleSize
                   c:separator c:extLst].freeze

      one_and_only_one "c:idx"
      zero_or_one "c:tx", successors: TAG_SEQ[3..]
      zero_or_one "c:spPr", successors: TAG_SEQ[5..]
      zero_or_one "c:txPr", successors: TAG_SEQ[6..]
      zero_or_one "c:dLblPos", successors: TAG_SEQ[7..]

      DEFAULT_XML = <<~XML.freeze
        <c:dLbl #{Ns.nsdecls("c", "a")}>
          <c:idx val="666"/>
          <c:spPr/>
          <c:txPr>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
              <a:pPr>
                <a:defRPr/>
              </a:pPr>
            </a:p>
          </c:txPr>
          <c:showLegendKey val="0"/>
          <c:showVal val="1"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
        </c:dLbl>
      XML

      def self.new_dLbl(context)
        context.build_from_xml(DEFAULT_XML)
      end

      # The label's own text replaces the generated one. A `c:spPr` or
      # `c:txPr` alongside `c:tx` makes a bubble chart unsaveable in
      # PowerPoint, so both go -- python-pptx does the same.
      def get_or_add_rich
        remove_spPr
        remove_txPr
        tx = get_or_add_tx
        tx.remove_strRef
        tx.get_or_add_rich
      end

      def rich
        xpath("./c:tx/c:rich").first
      end

      def remove_tx_rich
        tx = xpath("./c:tx[c:rich]").first
        remove(tx) if tx
      end
    end

    # -------------------------------------------------------------------------
    # Series, points and markers
    # -------------------------------------------------------------------------

    # `c:marker`, the symbol drawn at each point of a line, radar or XY series.
    class CT_Marker < Element
      tag "c:marker"
      zero_or_one "c:symbol", successors: %w[c:size c:spPr c:extLst]
      zero_or_one "c:size", successors: %w[c:spPr c:extLst]
      zero_or_one "c:spPr", successors: %w[c:extLst]
    end

    # `c:dPt`, the formatting of one data point.
    class CT_DPt < Element
      tag "c:dPt"
      TAG_SEQ = %w[c:idx c:invertIfNegative c:marker c:bubble3D c:explosion c:spPr
                   c:pictureOptions c:extLst].freeze

      one_and_only_one "c:idx"
      zero_or_one "c:marker", successors: TAG_SEQ[3..]
      zero_or_one "c:spPr", successors: TAG_SEQ[6..]

      def self.new_dPt(context)
        context.build_from_xml(%(<c:dPt #{Ns.nsdecls("c")}><c:idx val="0"/></c:dPt>))
      end
    end

    # `c:ser`, one series of a plot.
    #
    # Each plot family allows a different subset of these children, in this
    # relative order; python-pptx's superset sequence is used as-is so an
    # inserted element lands where it does there.
    class CT_Series < Element
      tag "c:ser"
      TAG_SEQ = %w[c:idx c:order c:tx c:spPr c:invertIfNegative c:pictureOptions c:marker
                   c:explosion c:dPt c:dLbls c:trendline c:errBars c:cat c:val c:xVal c:yVal
                   c:shape c:smooth c:bubbleSize c:bubble3D c:extLst].freeze

      zero_or_one "c:idx", successors: TAG_SEQ[1..]
      zero_or_one "c:order", successors: TAG_SEQ[2..]
      zero_or_one "c:tx", successors: TAG_SEQ[3..]
      zero_or_one "c:spPr", successors: TAG_SEQ[4..]
      zero_or_one "c:invertIfNegative", successors: TAG_SEQ[5..]
      zero_or_one "c:marker", successors: TAG_SEQ[7..]
      zero_or_more "c:dPt", successors: TAG_SEQ[9..], as: :dPt
      zero_or_one "c:dLbls", successors: TAG_SEQ[10..]
      zero_or_one "c:cat", successors: TAG_SEQ[13..]
      zero_or_one "c:val", successors: TAG_SEQ[14..]
      zero_or_one "c:xVal", successors: TAG_SEQ[15..]
      zero_or_one "c:yVal", successors: TAG_SEQ[16..]
      zero_or_one "c:smooth", successors: TAG_SEQ[18..]
      zero_or_one "c:bubbleSize", successors: TAG_SEQ[19..]

      def index
        idx&.val.to_i
      end

      def name
        xpath("./c:tx//c:pt/c:v").first&.text.to_s
      end

      def new_dLbls
        CT_DataLabels.new_data_labels(self)
      end

      def new_dPt
        CT_DPt.new_dPt(self)
      end

      %w[cat xVal yVal bubbleSize].each do |source|
        define_method(:"#{source}_ptCount_val") do
          xpath("./c:#{source}//c:ptCount/@val").first&.value.to_i
        end
      end

      # The label of the point at +idx+, or nil.
      def dLbl_for_point(idx)
        dLbls&.dLbl_for_point(idx)
      end

      def get_or_add_dLbl_for_point(idx)
        get_or_add_dLbls.get_or_add_dLbl_for_point(idx)
      end

      # The formatting record for the point at +idx+, created if absent.
      def get_or_add_dPt_for_point(idx)
        dPt_list.find { |point| point.idx.val == idx } || add_dPt.tap { |point| point.idx.val = idx }
      end
    end

    # A series' category, value and XY sources: a reference into the
    # worksheet plus the cached values.
    class CT_SeriesData < Element
      tag "c:cat", "c:val", "c:xVal", "c:yVal", "c:bubbleSize"
      zero_or_one "c:multiLvlStrRef", successors: []

      # The levels of a multi-level category source, leaf level first.
      def lvls
        xpath(".//c:lvl")
      end
    end

    # `c:pt`, one cached value in a series source.
    class CT_ChartPoint < Element
      tag "c:pt"
      required_attr "idx", type: SimpleTypes::XsdUnsignedInt

      def v
        xpath("./c:v").first
      end

      def value
        Float(v.text)
      end
    end

    # -------------------------------------------------------------------------
    # Plots
    # -------------------------------------------------------------------------

    # A plot element -- `c:barChart` and its siblings.
    #
    # They differ in which options they accept, but share enough that one
    # class covers what this library reads and writes.
    class CT_Plot < Element
      tag "c:barChart", "c:lineChart", "c:pieChart", "c:doughnutChart",
          "c:areaChart", "c:radarChart", "c:scatterChart", "c:bubbleChart", "c:area3DChart"

      zero_or_one "c:barDir", successors: %w[c:grouping c:varyColors c:ser c:dLbls
                                             c:gapWidth c:overlap c:serLines c:axId c:extLst]
      zero_or_one "c:grouping", successors: %w[c:varyColors c:ser c:dLbls c:gapWidth
                                               c:overlap c:serLines c:axId c:extLst]
      zero_or_one "c:varyColors", successors: %w[c:ser c:dLbls c:gapWidth c:overlap
                                                 c:serLines c:axId c:extLst]
      zero_or_one "c:dLbls", successors: %w[c:dropLines c:gapWidth c:overlap c:serLines
                                            c:bubble3D c:bubbleScale c:axId c:extLst]
      zero_or_one "c:gapWidth", successors: %w[c:overlap c:serLines c:axId c:extLst]
      zero_or_one "c:overlap", successors: %w[c:serLines c:axId c:extLst]
      zero_or_one "c:bubbleScale", successors: %w[c:showNegBubbles c:sizeRepresents c:axId
                                                  c:extLst]

      # Series in document order.
      def ser_list
        find_all("c:ser")
      end

      # Series in the order the chart draws them, which is `c:order`.
      def sers
        ser_list.sort_by { |ser| ser.order&.val.to_i }
      end

      def grouping_val
        grouping&.val || SimpleTypes::ST_Grouping::STANDARD
      end

      def new_dLbls
        CT_DataLabels.new_data_labels(self)
      end

      # The data-label settings, created with PowerPoint's defaults if the
      # plot has none yet.
      def get_or_add_default_dLbls
        dLbls || begin
          created = new_dLbls
          insert_dLbls(created)
          created
        end
      end

      # The first series' category source, which the others share.
      def cat
        xpath("./c:ser[1]/c:cat").first
      end

      def cat_pt_count
        xpath("./c:ser//c:cat//c:ptCount").first&.val.to_i
      end

      # One entry per category, nil where the workbook cell is empty -- such
      # a category has no `c:pt` but still counts in `c:ptCount`.
      def cat_pts
        points = xpath("./c:ser[1]/c:cat//c:lvl[1]/c:pt")
        points = xpath("./c:ser[1]/c:cat//c:pt") if points.empty?
        by_idx = points.to_h { |point| [point.idx, point] }
        Array.new(cat_pt_count) { |i| by_idx[i] }
      end
    end

    # -------------------------------------------------------------------------
    # Plot area, chart and chart space
    # -------------------------------------------------------------------------

    class CT_PlotArea < Element
      tag "c:plotArea"

      def plot_elements
        @node.element_children
             .select { |child| child.name.end_with?("Chart") }
             .map { |child| Element.wrap(child) }
      end

      def category_axis
        find("c:catAx") || find("c:dateAx")
      end

      def value_axes
        find_all("c:valAx")
      end

      # Every series in the chart, in plot order then series order -- which is
      # the order their chart-wide indices follow.
      def series_elements
        plot_elements.flat_map(&:ser_list)
      end

      def last_plot_element
        plot_elements.last
      end
    end

    # `c:chart`, which OOXML uses for two different things under one tag name.
    #
    # Inside `a:graphicData` it is a *reference*: an empty element carrying the
    # relationship id of the chart part. Inside `c:chartSpace` it is the chart
    # itself, with a title, plot area and legend. The registry is keyed by tag,
    # so one class covers both; an element only ever uses one half.
    class CT_Chart < Element
      tag "c:chart"

      # -- used in the `a:graphicData` reference form --
      optional_attr "r:id", type: SimpleTypes::ST_RelationshipId, as: :rId

      CHART_NS =
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' \
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

      def self.new_chart(r_id)
        Element.parse(%(<c:chart #{CHART_NS} r:id="#{r_id}"/>))
      end

      # -- used in the `c:chartSpace` form --
      TAG_SEQ = %w[c:title c:autoTitleDeleted c:pivotFmts c:view3D c:floor c:sideWall
                   c:backWall c:plotArea c:legend c:plotVisOnly c:dispBlanksAs
                   c:showDLblsOverMax c:extLst].freeze

      zero_or_one "c:title", successors: TAG_SEQ[1..]
      zero_or_one "c:autoTitleDeleted", successors: TAG_SEQ[2..]
      one_and_only_one "c:plotArea"
      zero_or_one "c:legend", successors: TAG_SEQ[9..]

      def new_title
        CT_Title.new_title(self)
      end
    end

    # `c:chartSpace`, the root of a chart part.
    class CT_ChartSpace < Element
      include ChartTextProperties

      tag "c:chartSpace"
      TAG_SEQ = %w[c:date1904 c:lang c:roundedCorners c:style c:clrMapOvr
                   c:pivotSource c:protection c:chart c:spPr c:txPr
                   c:externalData c:printSettings c:userShapes c:extLst].freeze

      zero_or_one "c:date1904", successors: TAG_SEQ[1..]
      zero_or_one "c:style", successors: TAG_SEQ[4..]
      one_and_only_one "c:chart"
      zero_or_one "c:txPr", successors: TAG_SEQ[10..]
      zero_or_one "c:externalData", successors: TAG_SEQ[11..]

      # The relationship id of the embedded workbook, or nil when the chart
      # has none.
      def xlsx_part_rId
        externalData&.rId
      end

      # The plot element, e.g. `c:barChart`, that this chart draws with.
      def plot_element
        xpath("./c:chart/c:plotArea/*").find { |e| e.nsptag.end_with?("Chart") }
      end

      def series_elements
        xpath("./c:chart/c:plotArea/*/c:ser")
      end
    end

    # `c:externalData`, the link from a chart to its embedded workbook.
    class CT_ExternalData < Element
      tag "c:externalData"
      optional_attr "r:id", type: SimpleTypes::ST_RelationshipId, as: :rId
      zero_or_one "c:autoUpdate", successors: []

      # PowerPoint always writes the auto-update flag, explicitly off.
      def ensure_auto_update
        get_or_add_autoUpdate.set("val", "0")
        self
      end
    end
  end
end
