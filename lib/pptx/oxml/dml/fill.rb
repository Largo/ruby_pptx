# frozen_string_literal: true

require "pptx/oxml/dml/color"

module Pptx
  module Oxml
    # `a:solidFill`, a single flat colour.
    class CT_SolidColorFillProperties < Element
      include ColorChoice

      tag "a:solidFill"
    end

    # `a:noFill`, explicitly transparent.
    class CT_NoFillProperties < Element
      tag "a:noFill"
    end

    # `a:grpFill`, inheriting the group's fill.
    class CT_GroupFillProperties < Element
      tag "a:grpFill"
    end

    # `a:gs`, one stop in a gradient.
    class CT_GradientStop < Element
      include ColorChoice

      tag "a:gs"
      required_attr "pos", type: SimpleTypes::ST_PositiveFixedPercentage
    end

    # `a:gsLst`, the stops of a gradient.
    class CT_GradientStopList < Element
      tag "a:gsLst"
      one_or_more "a:gs", as: :gs
    end

    # `a:lin`, the direction of a linear gradient.
    class CT_LinearShadeProperties < Element
      tag "a:lin"
      optional_attr "ang", type: SimpleTypes::ST_PositiveFixedAngle
      optional_attr "scaled", type: SimpleTypes::XsdBoolean
    end

    # `a:gradFill`, a gradient fill.
    class CT_GradientFillProperties < Element
      tag "a:gradFill"
      zero_or_one "a:gsLst", successors: %w[a:lin a:path a:tileRect]
      zero_or_one "a:lin", successors: %w[a:path a:tileRect]
      zero_or_one "a:path", successors: %w[a:tileRect]

      # The default gradient PowerPoint writes: two stops of the same theme
      # colour, one lightened, angled across the shape.
      DEFAULT_XML = <<~XML.freeze
        <a:gradFill #{Ns.nsdecls("a")} rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="accent1">
                <a:lumMod val="110000"/>
                <a:satMod val="105000"/>
                <a:tint val="67000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="50000">
              <a:schemeClr val="accent1">
                <a:lumMod val="105000"/>
                <a:satMod val="103000"/>
                <a:tint val="73000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="accent1">
                <a:lumMod val="105000"/>
                <a:satMod val="109000"/>
                <a:tint val="81000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      XML

      def self.new_grad_fill(context) = context.build_from_xml(DEFAULT_XML)
    end

    # `a:blip`, the image reference inside a picture fill.
    class CT_Blip < Element
      tag "a:blip"
      optional_attr "r:embed", type: SimpleTypes::ST_RelationshipId, as: :embed
      zero_or_one "a:extLst", successors: []

      # The Office 2016 extension that carries an SVG beside the raster the
      # blip itself points at. Consumers that understand it draw the vector;
      # those that do not draw the fallback.
      SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"

      # Point this blip's SVG extension at the part related by +r_id+.
      def add_svg_blip(r_id)
        ext_list = get_or_add_extLst
        ext_list.append(ext_list.build_from_xml(<<~XML))
          <a:ext #{Ns.nsdecls("a")} uri="#{SVG_EXT_URI}">
            <asvg:svgBlip #{Ns.nsdecls("asvg", "r")} r:embed="#{r_id}"/>
          </a:ext>
        XML
        self
      end

      # The relationship id of the SVG, or nil when this is a plain raster.
      def svg_rId = xpath("./a:extLst/a:ext/asvg:svgBlip/@r:embed").first&.value
    end

    # `a:extLst` as it appears on a blip.
    class CT_BlipExtensionList < Element
      tag "a:extLst"
      zero_or_more "a:ext", as: :ext
    end

    # `a:ext`, one DrawingML extension.
    class CT_BlipExtension < Element
      tag "a:ext"
      required_attr "uri", type: SimpleTypes::XsdString
    end

    # `a:blipFill`, a picture fill.
    class CT_BlipFillProperties < Element
      tag "a:blipFill"
      zero_or_one "a:blip", successors: %w[a:srcRect a:tile a:stretch]
      zero_or_one "a:srcRect", successors: %w[a:tile a:stretch]
    end

    # `a:pattFill`, a two-colour pattern fill.
    class CT_PatternFillProperties < Element
      tag "a:pattFill"
      zero_or_one "a:fgClr", successors: %w[a:bgClr]
      zero_or_one "a:bgClr", successors: []
      optional_attr "prst", type: Enum::MSO_PATTERN_TYPE
    end
  end
end
