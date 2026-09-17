# frozen_string_literal: true

require "pptx/oxml/shapes/shared"
require "pptx/oxml/text"

module Pptx
  module Oxml
    # `a:avLst`, the adjustment values of a preset geometry.
    class CT_GeomGuideList < Element
      tag "a:avLst"
      zero_or_more "a:gd", as: :gd
    end

    # `a:gd`, one geometry guide (adjustment value).
    class CT_GeomGuide < Element
      tag "a:gd"
      required_attr "name", type: SimpleTypes::XsdString
      required_attr "fmla", type: SimpleTypes::XsdString
    end

    # `a:prstGeom`, a preset (built-in) shape geometry.
    class CT_PresetGeometry2D < Element
      tag "a:prstGeom"
      zero_or_one "a:avLst", successors: []
      required_attr "prst", type: Enum::MSO_AUTO_SHAPE_TYPE

      # The adjustment guides, in document order.
      def gd_list = avLst ? avLst.gd_list : []

      # Replace the adjustment values with +name_value_pairs+.
      def rewrite_guides(name_value_pairs)
        remove_avLst
        av_lst = get_or_add_avLst
        name_value_pairs.each do |name, value|
          av_lst.add_gd(name: name, fmla: "val #{value}")
        end
      end
    end

    # `a:custGeom`, a freeform shape geometry.
    class CT_CustomGeometry2D < Element
      tag "a:custGeom"
      zero_or_one "a:pathLst", successors: []
    end

    # `p:cNvSpPr`, the non-visual properties specific to a `p:sp`.
    class CT_NonVisualDrawingShapeProps < Element
      tag "p:cNvSpPr"
      zero_or_one "a:spLocks", successors: %w[a:extLst]
      optional_attr "txBox", type: SimpleTypes::XsdBoolean
    end

    # `p:nvSpPr`, the non-visual properties group of a `p:sp`.
    class CT_ShapeNonVisual < Element
      tag "p:nvSpPr"
      one_and_only_one "p:cNvPr"
      one_and_only_one "p:cNvSpPr"
      one_and_only_one "p:nvPr"
    end

    # `p:sp`, an auto shape, text box, placeholder or freeform.
    class CT_Shape < Element
      include BaseShapeElement
      tag "p:sp"
      one_and_only_one "p:nvSpPr"
      one_and_only_one "p:spPr"
      zero_or_one "p:txBody", successors: %w[p:extLst]

      # Placeholder types that come with a text frame when created.
      TEXTUAL_PLACEHOLDERS = [
        Enum::PP_PLACEHOLDER::TITLE, Enum::PP_PLACEHOLDER::CENTER_TITLE,
        Enum::PP_PLACEHOLDER::SUBTITLE, Enum::PP_PLACEHOLDER::BODY,
        Enum::PP_PLACEHOLDER::OBJECT
      ].freeze

      class << self
        # A `p:sp` configured as a placeholder inheriting from a layout.
        #
        # The shape carries no geometry or position of its own: a placeholder
        # inherits both from the corresponding placeholder on its layout.
        def new_placeholder_sp(id, name, ph_type, orient, sz, idx)
          sp = Element.parse(<<~XML)
            <p:sp #{Ns.nsdecls('a', 'p')}>
              <p:nvSpPr>
                <p:cNvPr id="#{id}" name="#{escape(name)}"/>
                <p:cNvSpPr>
                  <a:spLocks noGrp="1"/>
                </p:cNvSpPr>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr/>
            </p:sp>
          XML

          ph = sp.nvSpPr.nvPr.get_or_add_ph
          ph.type = ph_type
          ph.idx = idx
          ph.orient = orient
          ph.sz = sz

          sp.append(CT_TextBody.new_element(sp)) if TEXTUAL_PLACEHOLDERS.include?(ph_type)
          
          sp
        end

        # A `p:sp` configured as an auto shape with preset geometry.
        def new_autoshape_sp(id, name, prst, x, y, cx, cy)
          Element.parse(<<~XML)
            <p:sp #{Ns.nsdecls('a', 'p')}>
              <p:nvSpPr>
                <p:cNvPr id="#{id}" name="#{escape(name)}"/>
                <p:cNvSpPr/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="#{x.to_i}" y="#{y.to_i}"/>
                  <a:ext cx="#{cx.to_i}" cy="#{cy.to_i}"/>
                </a:xfrm>
                <a:prstGeom prst="#{prst}">
                  <a:avLst/>
                </a:prstGeom>
              </p:spPr>
              #{DEFAULT_STYLE}
              <p:txBody>
                <a:bodyPr rtlCol="0" anchor="ctr"/>
                <a:lstStyle/>
                <a:p>
                  <a:pPr algn="ctr"/>
                </a:p>
              </p:txBody>
            </p:sp>
          XML
        end

        # A `p:sp` configured as a text box.
        def new_textbox_sp(id, name, x, y, cx, cy)
          Element.parse(<<~XML)
            <p:sp #{Ns.nsdecls('a', 'p')}>
              <p:nvSpPr>
                <p:cNvPr id="#{id}" name="#{escape(name)}"/>
                <p:cNvSpPr txBox="1"/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="#{x.to_i}" y="#{y.to_i}"/>
                  <a:ext cx="#{cx.to_i}" cy="#{cy.to_i}"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                  <a:avLst/>
                </a:prstGeom>
                <a:noFill/>
              </p:spPr>
              <p:txBody>
                <a:bodyPr wrap="none">
                  <a:spAutoFit/>
                </a:bodyPr>
                <a:lstStyle/>
                <a:p/>
              </p:txBody>
            </p:sp>
          XML
        end

        def escape(text) = text.to_s.gsub("&", "&amp;").gsub("<", "&lt;").gsub('"', "&quot;")

        DEFAULT_STYLE = <<~XML
          <p:style>
            <a:lnRef idx="1">
              <a:schemeClr val="accent1"/>
            </a:lnRef>
            <a:fillRef idx="3">
              <a:schemeClr val="accent1"/>
            </a:fillRef>
            <a:effectRef idx="2">
              <a:schemeClr val="accent1"/>
            </a:effectRef>
            <a:fontRef idx="minor">
              <a:schemeClr val="lt1"/>
            </a:fontRef>
          </p:style>
        XML
      end

      # A text body added to a shape that has none needs the minimum structure
      # the schema requires, not an empty element.
      def new_txBody = CT_TextBody.new_element(self)

      def prstGeom = spPr.prstGeom

      def prst = prstGeom&.prst

      def ln = spPr.ln

      def get_or_add_ln = spPr.get_or_add_ln

      # A shape has custom geometry -- is a freeform -- when it carries
      # `a:custGeom` in place of `a:prstGeom`.
      def custom_geometry? = !spPr.custGeom.nil?

      # A text box is an `p:sp` whose `p:cNvSpPr` says `txBox="1"`.
      def textbox? = nvSpPr.cNvSpPr.txBox == true

      # An auto shape has preset geometry and is not a text box.
      def autoshape? = !prstGeom.nil? && nvSpPr.cNvSpPr.txBox != true
    end

  end
end
