# frozen_string_literal: true

require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/oxml/simple_types"

module Pptx
  module Oxml
    # `c:chart` as it appears inside a `p:graphicFrame`, pointing at the chart
    # part by relationship id.
    class CT_GraphicFrameChart < Element
      tag "c:chart"
      optional_attr "r:id", type: SimpleTypes::ST_RelationshipId, as: :rId

      CHART_NS =
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' \
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

      def self.new_chart(r_id)
        Element.parse(%(<c:chart #{CHART_NS} r:id="#{r_id}"/>))
      end
    end

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
