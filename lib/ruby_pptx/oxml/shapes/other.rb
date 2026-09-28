# frozen_string_literal: true

require "ruby_pptx/oxml/shapes/shared"
require "ruby_pptx/oxml/table"
require "ruby_pptx/oxml/chart"

module Pptx
  module Oxml
    # The non-visual property groups of the non-`p:sp` shapes. They differ in
    # what else they hold, but each starts with the `p:cNvPr` that carries the
    # shape's id and name, which is all this library needs from them.
    class CT_ShapeNonVisualCommon < Element
      tag "p:nvPicPr", "p:nvGraphicFramePr"
      one_and_only_one "p:cNvPr"
      zero_or_one "p:nvPr", successors: []
    end

    # `p:pic`, a picture (or a movie, which is a picture with media relations).
    class CT_Picture < Element
      include BaseShapeElement

      tag "p:pic"
      one_and_only_one "p:nvPicPr"
      one_and_only_one "p:blipFill"
      one_and_only_one "p:spPr"

      # The `a:blip` inside this picture's fill.
      def blip
        xpath("./p:blipFill/a:blip").first
      end

      # The relationship id of the image this picture shows, or nil.
      def blip_rId
        blip&.embed
      end

      # Cropping, as a fraction of the image cropped from each side. Reading
      # gives 0.0 when there is no `a:srcRect`; writing creates one.
      %w[l t r b].each do |side|
        define_method(:"srcRect_#{side}") { blipFill.srcRect&.public_send(side) || 0.0 }
        define_method(:"srcRect_#{side}=") do |value|
          blipFill.get_or_add_srcRect.public_send(:"#{side}=", value)
        end
      end

      def ln
        spPr.ln
      end

      def get_or_add_ln
        spPr.get_or_add_ln
      end

      # Crop so an image of +image_size+ fills +view_size+ exactly when
      # stretched with its aspect ratio kept: the excess is cut equally from
      # both ends of whichever dimension is too long. Both sizes are
      # (width, height); only their ratios matter, so the units need not agree.
      def crop_to_fit(image_size, view_size)
        left, top, right, bottom = fill_cropping(image_size, view_size)
        rect = blipFill.get_or_add_srcRect
        rect.l = left
        rect.t = top
        rect.r = right
        rect.b = bottom
      end

      # A picture placeholder once filled: a `p:pic` with no `a:xfrm`, so its
      # position and size keep coming from the layout.
      def self.new_ph_pic(id, name, desc, r_id)
        Element.parse(<<~XML)
          <p:pic #{Ns.nsdecls("p", "a", "r")}>
            <p:nvPicPr>
              <p:cNvPr id="#{id}" name="#{escape(name)}" descr="#{escape(desc)}"/>
              <p:cNvPicPr>
                <a:picLocks noGrp="1" noChangeAspect="1"/>
              </p:cNvPicPr>
              <p:nvPr/>
            </p:nvPicPr>
            <p:blipFill>
              <a:blip r:embed="#{r_id}"/>
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </p:blipFill>
            <p:spPr/>
          </p:pic>
        XML
      end

      # A movie is a `p:pic` whose non-visual properties name a video file.
      def movie?
        !xpath("./p:nvPicPr/p:nvPr/a:videoFile").empty?
      end

      # A `p:pic` displaying the image related by +r_id+.
      def self.new_pic(id, name, desc, r_id, x, y, cx, cy)
        Element.parse(<<~XML)
          <p:pic #{Ns.nsdecls("a", "p", "r")}>
            <p:nvPicPr>
              <p:cNvPr id="#{id}" name="#{escape(name)}" descr="#{escape(desc)}"/>
              <p:cNvPicPr>
                <a:picLocks noChangeAspect="1"/>
              </p:cNvPicPr>
              <p:nvPr/>
            </p:nvPicPr>
            <p:blipFill>
              <a:blip r:embed="#{r_id}"/>
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </p:blipFill>
            <p:spPr>
              <a:xfrm>
                <a:off x="#{x.to_i}" y="#{y.to_i}"/>
                <a:ext cx="#{cx.to_i}" cy="#{cy.to_i}"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </p:spPr>
          </p:pic>
        XML
      end

      # A `p:pic` showing a video.
      #
      # The video is referenced three times over: as a legacy `a:videoFile`
      # link, as a `p14:media` embed for PowerPoint 2010 and later, and as the
      # `ppaction://media` click action. The blip points at a poster frame,
      # which is the still PowerPoint shows before the video plays.
      def self.new_video_pic(id, name, video_r_id, media_r_id, poster_r_id, x, y, cx, cy)
        Element.parse(<<~XML)
          <p:pic #{Ns.nsdecls("a", "p", "r")}>
            <p:nvPicPr>
              <p:cNvPr id="#{id}" name="#{escape(name)}">
                <a:hlinkClick r:id="" action="ppaction://media"/>
              </p:cNvPr>
              <p:cNvPicPr>
                <a:picLocks noChangeAspect="1"/>
              </p:cNvPicPr>
              <p:nvPr>
                <a:videoFile r:link="#{video_r_id}"/>
                <p:extLst>
                  <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
                    <p14:media #{Ns.nsdecls("p14")} r:embed="#{media_r_id}"/>
                  </p:ext>
                </p:extLst>
              </p:nvPr>
            </p:nvPicPr>
            <p:blipFill>
              <a:blip r:embed="#{poster_r_id}"/>
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </p:blipFill>
            <p:spPr>
              <a:xfrm>
                <a:off x="#{x.to_i}" y="#{y.to_i}"/>
                <a:ext cx="#{cx.to_i}" cy="#{cy.to_i}"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </p:spPr>
          </p:pic>
        XML
      end

      def self.escape(text)
        text.to_s.gsub("&", "&amp;").gsub("<", "&lt;").gsub(">", "&gt;").gsub('"', "&quot;")
      end

      private

      # (left, top, right, bottom) fractions to crop. The arithmetic is
      # python-pptx's, step for step, so the percentages round identically.
      def fill_cropping(image_size, view_size)
        view_ratio = view_size[0].to_f / view_size[1]
        image_ratio = image_size[0].to_f / image_size[1]
        if view_ratio < image_ratio # image too wide
          crop = (1.0 - (view_ratio / image_ratio)) / 2.0
          [crop, 0.0, crop, 0.0]
        elsif view_ratio > image_ratio # image too tall
          crop = (1.0 - (image_ratio / view_ratio)) / 2.0
          [0.0, crop, 0.0, crop]
        else
          [0.0, 0.0, 0.0, 0.0]
        end
      end
    end

    # `p:nvCxnSpPr`, a connector's non-visual properties.
    class CT_ConnectorNonVisual < Element
      tag "p:nvCxnSpPr"
      one_and_only_one "p:cNvPr"
      one_and_only_one "p:cNvCxnSpPr"
      zero_or_one "p:nvPr", successors: []
    end

    # `p:cNvCxnSpPr`, which records what each end of a connector attaches to.
    class CT_NonVisualConnectorProperties < Element
      tag "p:cNvCxnSpPr"
      zero_or_one "a:stCxn", successors: %w[a:endCxn a:extLst]
      zero_or_one "a:endCxn", successors: %w[a:extLst]
    end

    # `a:stCxn` and `a:endCxn`: one end of a connector attached to a shape's
    # connection point.
    class CT_Connection < Element
      tag "a:stCxn", "a:endCxn"
      required_attr "id", type: SimpleTypes::ST_DrawingElementId
      required_attr "idx", type: SimpleTypes::XsdUnsignedInt
    end

    # `p:cxnSp`, a connector.
    class CT_Connector < Element
      include BaseShapeElement

      tag "p:cxnSp"
      one_and_only_one "p:nvCxnSpPr"
      one_and_only_one "p:spPr"

      # A connector is stored as a bounding box plus flip flags rather than as
      # two points, so a line running right-to-left is the same box with
      # flipH set.
      def self.new_cxnSp(id, name, prst, x, y, cx, cy, flip_h, flip_v)
        flip = +""
        flip << %( flipH="1") if flip_h
        flip << %( flipV="1") if flip_v

        Element.parse(<<~XML)
          <p:cxnSp #{Ns.nsdecls("a", "p")}>
            <p:nvCxnSpPr>
              <p:cNvPr id="#{id}" name="#{CT_Picture.escape(name)}"/>
              <p:cNvCxnSpPr/>
              <p:nvPr/>
            </p:nvCxnSpPr>
            <p:spPr>
              <a:xfrm#{flip}>
                <a:off x="#{x.to_i}" y="#{y.to_i}"/>
                <a:ext cx="#{cx.to_i}" cy="#{cy.to_i}"/>
              </a:xfrm>
              <a:prstGeom prst="#{prst}">
                <a:avLst/>
              </a:prstGeom>
            </p:spPr>
            <p:style>
              <a:lnRef idx="2">
                <a:schemeClr val="accent1"/>
              </a:lnRef>
              <a:fillRef idx="0">
                <a:schemeClr val="accent1"/>
              </a:fillRef>
              <a:effectRef idx="1">
                <a:schemeClr val="accent1"/>
              </a:effectRef>
              <a:fontRef idx="minor">
                <a:schemeClr val="tx1"/>
              </a:fontRef>
            </p:style>
          </p:cxnSp>
        XML
      end
    end

    # `p:oleObj`, an OLE object inside a graphic frame.
    class CT_OleObject < Element
      tag "p:oleObj"
      optional_attr "progId", type: SimpleTypes::XsdString
      optional_attr "r:id", type: SimpleTypes::XsdString, as: :rId
      optional_attr "showAsIcon", type: SimpleTypes::XsdBoolean, default: false
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
      def xfrm
        find("p:xfrm")
      end

      def get_or_add_xfrm
        xfrm
      end

      def graphic_data_uri
        xpath("./a:graphic/a:graphicData/@uri").first&.value
      end

      # The relationship id of the chart part this frame refers to.
      def chart_rId
        xpath("./a:graphic/a:graphicData/c:chart/@r:id").first&.value
      end

      def table?
        graphic_data_uri == URI_TABLE
      end

      def chart?
        graphic_data_uri == URI_CHART
      end

      # The `a:tbl` inside this frame, or nil when it holds something else.
      def tbl
        xpath("./a:graphic/a:graphicData/a:tbl").first
      end

      def ole_object?
        graphic_data_uri == URI_OLE_OBJECT
      end

      # The `p:oleObj` inside this frame, or nil.
      def oleObj
        xpath("./a:graphic/a:graphicData/p:oleObj").first
      end

      # An OLE object carries its file (`p:embed`) or points elsewhere for it.
      def embedded_ole_object?
        !xpath("./a:graphic/a:graphicData/p:oleObj/p:embed").empty?
      end

      class << self
        # A frame holding an OLE object shown as an icon: the embedded file is
        # related by +ole_r_id+, the icon image by +icon_r_id+. Written as
        # python-pptx writes it.
        def new_ole_object_graphic_frame(id, name, ole_r_id, prog_id, icon_r_id, x, y, cx, cy, img_w, img_h)
          Element.parse(<<~XML)
            <p:graphicFrame #{Ns.nsdecls("a", "p", "r")}>
              <p:nvGraphicFramePr>
                <p:cNvPr id="#{id}" name="#{CT_Picture.escape(name)}"/>
                <p:cNvGraphicFramePr>
                  <a:graphicFrameLocks noGrp="1"/>
                </p:cNvGraphicFramePr>
                <p:nvPr/>
              </p:nvGraphicFramePr>
              <p:xfrm>
                <a:off x="#{x.to_i}" y="#{y.to_i}"/>
                <a:ext cx="#{cx.to_i}" cy="#{cy.to_i}"/>
              </p:xfrm>
              <a:graphic>
                <a:graphicData uri="#{URI_OLE_OBJECT}">
                  <p:oleObj showAsIcon="1" r:id="#{ole_r_id}" imgW="#{img_w.to_i}" imgH="#{img_h.to_i}" progId="#{CT_Picture.escape(prog_id)}">
                    <p:embed/>
                    <p:pic>
                      <p:nvPicPr>
                        <p:cNvPr id="0" name=""/>
                        <p:cNvPicPr/>
                        <p:nvPr/>
                      </p:nvPicPr>
                      <p:blipFill>
                        <a:blip r:embed="#{icon_r_id}"/>
                        <a:stretch>
                          <a:fillRect/>
                        </a:stretch>
                      </p:blipFill>
                      <p:spPr>
                        <a:xfrm>
                          <a:off x="#{x.to_i}" y="#{y.to_i}"/>
                          <a:ext cx="#{cx.to_i}" cy="#{cy.to_i}"/>
                        </a:xfrm>
                        <a:prstGeom prst="rect">
                          <a:avLst/>
                        </a:prstGeom>
                      </p:spPr>
                    </p:pic>
                  </p:oleObj>
                </a:graphicData>
              </a:graphic>
            </p:graphicFrame>
          XML
        end

        # An empty `p:graphicFrame`. It is not a valid shape until a graphical
        # object such as a table is placed inside it.
        def new_graphic_frame(id, name, x, y, cx, cy)
          Element.parse(<<~XML)
            <p:graphicFrame #{Ns.nsdecls("a", "p")}>
              <p:nvGraphicFramePr>
                <p:cNvPr id="#{id}" name="#{CT_Picture.escape(name)}"/>
                <p:cNvGraphicFramePr>
                  <a:graphicFrameLocks noGrp="1"/>
                </p:cNvGraphicFramePr>
                <p:nvPr/>
              </p:nvGraphicFramePr>
              <p:xfrm>
                <a:off x="#{x.to_i}" y="#{y.to_i}"/>
                <a:ext cx="#{cx.to_i}" cy="#{cy.to_i}"/>
              </p:xfrm>
              <a:graphic>
                <a:graphicData/>
              </a:graphic>
            </p:graphicFrame>
          XML
        end

        # A `p:graphicFrame` referring to the chart part related by +r_id+.
        def new_chart_graphic_frame(id, name, r_id, x, y, cx, cy)
          frame = new_graphic_frame(id, name, x, y, cx, cy)
          graphic_data = frame.xpath("./a:graphic/a:graphicData").first
          graphic_data.set("uri", URI_CHART)
          graphic_data.append(
            graphic_data.import(CT_Chart.new_chart(r_id))
          )
          frame
        end

        # A `p:graphicFrame` containing a table of +rows+ by +cols+.
        def new_table_graphic_frame(id, name, rows, cols, x, y, cx, cy)
          frame = new_graphic_frame(id, name, x, y, cx, cy)
          graphic_data = frame.xpath("./a:graphic/a:graphicData").first
          graphic_data.set("uri", URI_TABLE)
          graphic_data.append(CT_Table.new_tbl(graphic_data, rows, cols, cx, cy))
          frame
        end
      end
    end

    # `p:contentPart`, which this library carries through without modelling.
    class CT_ContentPart < Element
      include BaseShapeElement

      tag "p:contentPart"
    end
  end
end
