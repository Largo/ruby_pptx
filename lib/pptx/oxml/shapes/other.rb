# frozen_string_literal: true

require "pptx/oxml/shapes/shared"

module Pptx
  module Oxml
    # `p:pic`, a picture (or a movie, which is a picture with media relations).
    class CT_Picture < Element
      include BaseShapeElement
      tag "p:pic"
      one_and_only_one "p:nvPicPr"
      one_and_only_one "p:blipFill"
      one_and_only_one "p:spPr"

      # A movie is a `p:pic` whose non-visual properties name a video file.
      def movie? = !xpath("./p:nvPicPr/p:nvPr/a:videoFile").empty?
    end

    # `p:cxnSp`, a connector.
    class CT_Connector < Element
      include BaseShapeElement
      tag "p:cxnSp"
      one_and_only_one "p:nvCxnSpPr"
      one_and_only_one "p:spPr"
    end

    # `p:graphicFrame`, the container for a table, chart or embedded object.
    class CT_GraphicalObjectFrame < Element
      include BaseShapeElement
      tag "p:graphicFrame"
      one_and_only_one "p:nvGraphicFramePr"
      one_and_only_one "p:xfrm"
      one_and_only_one "a:graphic"

      URI_TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table"
      URI_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
      URI_OLE_OBJECT =
        "http://schemas.openxmlformats.org/presentationml/2006/ole"

      # A graphic frame keeps its transform directly, not under `p:spPr`.
      def xfrm = find("p:xfrm")

      def get_or_add_xfrm = xfrm

      def graphic_data_uri = xpath("./a:graphic/a:graphicData/@uri").first&.value

      def table? = graphic_data_uri == URI_TABLE
      def chart? = graphic_data_uri == URI_CHART
    end

    # `p:contentPart`, which this library carries through without modelling.
    class CT_ContentPart < Element
      include BaseShapeElement
      tag "p:contentPart"
    end
  end
end
