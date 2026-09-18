# frozen_string_literal: true

require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/oxml/simple_types"

module Pptx
  module Oxml
    # `c:autoUpdate`, whether the chart refreshes itself from the workbook.
    class CT_AutoUpdate < Element
      tag "c:autoUpdate"
      optional_attr "val", type: SimpleTypes::XsdBoolean
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

    # A boolean-valued `val` element, e.g. `c:delete` or `c:showVal`.
    class CT_ChartBoolean < Element
      tag "c:delete", "c:overlay", "c:showVal", "c:showCatName", "c:showSerName",
          "c:showPercent", "c:showLegendKey", "c:showBubbleSize", "c:varyColors",
          "c:autoTitleDeleted", "c:plotVisOnly", "c:showDLblsOverMax",
          "c:showLeaderLines"
      optional_attr "val", type: SimpleTypes::XsdBoolean
    end

    # A double-valued `val` element, e.g. `c:max` on an axis scale.
    class CT_ChartDouble < Element
      tag "c:max", "c:min", "c:majorUnit", "c:minorUnit"
      required_attr "val", type: SimpleTypes::XsdDouble
    end

    # A string-valued `val` element, e.g. `c:legendPos` or `c:tickLblPos`.
    class CT_ChartString < Element
      tag "c:legendPos", "c:tickLblPos", "c:dLblPos", "c:orientation",
          "c:majorTickMark", "c:minorTickMark", "c:crosses", "c:axPos",
          "c:barDir", "c:grouping", "c:radarStyle", "c:scatterStyle"
      optional_attr "val", type: SimpleTypes::XsdString
    end

    # `c:gapWidth` and `c:overlap`, which are percentages stored as integers.
    class CT_ChartPercent < Element
      tag "c:gapWidth", "c:overlap"
      optional_attr "val", type: SimpleTypes::XsdInt
    end

    # `c:numFmt`, a number format applied to an axis or to data labels.
    class CT_NumberFormat < Element
      tag "c:numFmt"
      optional_attr "formatCode", type: SimpleTypes::XsdString
      optional_attr "sourceLinked", type: SimpleTypes::XsdBoolean
    end

    # `c:scaling`, the axis range.
    class CT_Scaling < Element
      tag "c:scaling"
      zero_or_one "c:logBase", successors: %w[c:orientation c:max c:min c:extLst]
      zero_or_one "c:orientation", successors: %w[c:max c:min c:extLst]
      zero_or_one "c:max", successors: %w[c:min c:extLst]
      zero_or_one "c:min", successors: %w[c:extLst]

      # nil means "auto", which is what an absent element says.
      def maximum = max&.val

      def maximum=(value)
        value.nil? ? remove_max : (get_or_add_max.val = value)
        value
      end

      def minimum = min&.val

      def minimum=(value)
        value.nil? ? remove_min : (get_or_add_min.val = value)
        value
      end
    end

    # `c:catAx` and `c:valAx`, sharing enough structure to be one class.
    class CT_Axis < Element
      tag "c:catAx", "c:valAx", "c:dateAx"
      TAG_SEQ = %w[c:axId c:scaling c:delete c:axPos c:majorGridlines c:minorGridlines
                   c:title c:numFmt c:majorTickMark c:minorTickMark c:tickLblPos
                   c:spPr c:txPr c:crossAx c:crosses c:crossesAt c:crossBetween
                   c:majorUnit c:minorUnit c:dispUnits c:extLst].freeze

      one_and_only_one "c:scaling"
      zero_or_one "c:delete", successors: TAG_SEQ[3..]
      zero_or_one "c:majorGridlines", successors: TAG_SEQ[5..]
      zero_or_one "c:minorGridlines", successors: TAG_SEQ[6..]
      zero_or_one "c:title", successors: TAG_SEQ[7..]
      zero_or_one "c:numFmt", successors: TAG_SEQ[8..]
      zero_or_one "c:majorTickMark", successors: TAG_SEQ[9..]
      zero_or_one "c:minorTickMark", successors: TAG_SEQ[10..]
      zero_or_one "c:tickLblPos", successors: TAG_SEQ[11..]
      zero_or_one "c:majorUnit", successors: TAG_SEQ[18..]
      zero_or_one "c:minorUnit", successors: TAG_SEQ[19..]
    end

    # `c:majorGridlines` and `c:minorGridlines`.
    class CT_ChartLines < Element
      tag "c:majorGridlines", "c:minorGridlines"
    end

    # `c:legend`.
    class CT_Legend < Element
      tag "c:legend"
      TAG_SEQ = %w[c:legendPos c:legendEntry c:layout c:overlay c:spPr c:txPr c:extLst].freeze

      zero_or_one "c:legendPos", successors: TAG_SEQ[1..]
      zero_or_one "c:layout", successors: TAG_SEQ[3..]
      zero_or_one "c:overlay", successors: TAG_SEQ[4..]
      zero_or_one "c:txPr", successors: TAG_SEQ[6..]
    end

    # `c:title`, the chart or axis title.
    class CT_Title < Element
      tag "c:title"
      TAG_SEQ = %w[c:tx c:layout c:overlay c:spPr c:txPr c:extLst].freeze

      zero_or_one "c:tx", successors: TAG_SEQ[1..]
      zero_or_one "c:layout", successors: TAG_SEQ[2..]
      zero_or_one "c:overlay", successors: TAG_SEQ[4..]

      # A title with a rich-text body, which is what PowerPoint writes when a
      # title is typed rather than taken from the data.
      def self.new_title(context)
        context.build_from_xml(<<~XML)
          <c:title #{Ns.nsdecls('c', 'a')}>
            <c:tx>
              <c:rich>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr>
                    <a:defRPr/>
                  </a:pPr>
                  <a:r>
                    <a:t/>
                  </a:r>
                </a:p>
              </c:rich>
            </c:tx>
            <c:layout/>
            <c:overlay val="0"/>
          </c:title>
        XML
      end

      # The `c:tx/c:rich` body, which is a text body like any other.
      def rich = xpath("./c:tx/c:rich").first
    end

    # `c:tx` on a title, holding either rich text or a cell reference.
    class CT_TitleText < Element
      tag "c:tx"
      zero_or_one "c:rich", successors: []
    end

    # `c:dLbls`, data-label settings for a plot or series.
    class CT_DataLabels < Element
      tag "c:dLbls"
      TAG_SEQ = %w[c:dLbl c:numFmt c:spPr c:txPr c:dLblPos c:showLegendKey c:showVal
                   c:showCatName c:showSerName c:showPercent c:showBubbleSize
                   c:separator c:showLeaderLines c:leaderLines c:extLst].freeze

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

      # The element PowerPoint writes when data labels are switched on: every
      # kind of label named and turned off, with leader lines allowed. An
      # empty `c:dLbls` would not satisfy the schema, which wants the group
      # present.
      DEFAULT_XML = <<~XML
        <c:dLbls #{Ns.nsdecls('c')}>
          <c:showLegendKey val="0"/>
          <c:showVal val="0"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
          <c:showLeaderLines val="1"/>
        </c:dLbls>
      XML

      def self.new_data_labels(context) = context.build_from_xml(DEFAULT_XML)
    end

    # A plot element -- `c:barChart` and its siblings.
    #
    # They differ in which options they accept, but share enough that one
    # class covers what this library reads and writes.
    class CT_Plot < Element
      tag "c:barChart", "c:lineChart", "c:pieChart", "c:doughnutChart",
          "c:areaChart", "c:radarChart", "c:scatterChart", "c:bubbleChart"

      zero_or_one "c:varyColors", successors: %w[c:ser c:dLbls c:gapWidth c:overlap
                                                 c:serLines c:axId c:extLst]
      zero_or_one "c:dLbls", successors: %w[c:gapWidth c:overlap c:serLines c:axId c:extLst]
      zero_or_one "c:gapWidth", successors: %w[c:overlap c:serLines c:axId c:extLst]
      zero_or_one "c:overlap", successors: %w[c:serLines c:axId c:extLst]

      def ser_list = find_all("c:ser")

      # The data-label settings, created with PowerPoint's defaults if the
      # plot has none yet.
      def get_or_add_default_dLbls
        dLbls || begin
          created = CT_DataLabels.new_data_labels(self)
          insert_dLbls(created)
          created
        end
      end
    end

    # `c:ser`, one series of a plot.
    #
    # The families order their series children slightly differently -- a line
    # series has `c:marker` and `c:smooth` where a bar series has
    # `c:invertIfNegative` and `c:shape`. The successor lists below are a
    # superset covering all of them, so an element inserted here lands in the
    # right place whichever family it belongs to.
    class CT_Series < Element
      tag "c:ser"
      TAG_SEQ = %w[c:idx c:order c:tx c:spPr c:invertIfNegative c:marker
                   c:pictureOptions c:dPt c:dLbls c:trendline c:errBars c:cat
                   c:xVal c:yVal c:val c:bubbleSize c:bubble3D c:shape c:smooth
                   c:extLst].freeze

      zero_or_one "c:idx", successors: TAG_SEQ[1..]
      zero_or_one "c:order", successors: TAG_SEQ[2..]
      zero_or_one "c:tx", successors: TAG_SEQ[3..]
      zero_or_one "c:spPr", successors: TAG_SEQ[4..]
      zero_or_one "c:dLbls", successors: TAG_SEQ[9..]
      zero_or_one "c:cat", successors: TAG_SEQ[12..]
      zero_or_one "c:val", successors: TAG_SEQ[15..]

      def index = idx&.val.to_i

      def name = xpath("./c:tx//c:pt/c:v").first&.text.to_s
    end

    # `c:idx` and `c:order`, a series' position in the chart.
    class CT_SeriesIndex < Element
      tag "c:idx", "c:order"
      required_attr "val", type: SimpleTypes::XsdUnsignedInt
    end

    # `c:tx`, `c:cat`, `c:val` and their XY counterparts: a reference into the
    # worksheet plus the cached values.
    class CT_SeriesData < Element
      tag "c:tx", "c:cat", "c:val", "c:xVal", "c:yVal", "c:bubbleSize"
    end

    # `c:plotArea`.
    class CT_PlotArea < Element
      tag "c:plotArea"

      def plot_elements
        @node.element_children
             .select { |child| child.name.end_with?("Chart") }
             .map { |child| Element.wrap(child) }
      end

      def category_axis = find("c:catAx") || find("c:dateAx")

      def value_axis = find("c:valAx")

      def value_axes = find_all("c:valAx")

      # Every series in the chart, in plot order then series order -- which is
      # the order their chart-wide indices follow.
      def series_elements = plot_elements.flat_map(&:ser_list)

      def last_plot_element = plot_elements.last
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
    end

    # `c:chartSpace`, the root of a chart part.
    class CT_ChartSpace < Element
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
      def xlsx_part_rId = externalData&.rId

      # The plot element, e.g. `c:barChart`, that this chart draws with.
      def plot_element
        xpath("./c:chart/c:plotArea/*").find { |e| e.nsptag.end_with?("Chart") }
      end

      def series_elements = xpath("./c:chart/c:plotArea/*/c:ser")
    end
  end
end
