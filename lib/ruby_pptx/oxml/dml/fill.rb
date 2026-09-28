# frozen_string_literal: true

require "ruby_pptx/oxml/dml/color"

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

      # python-pptx's default, verbatim: PowerPoint's "White" template
      # gradient, two accent-1 stops on a linear path. The `a:lin` carries no
      # angle, so the direction is inherited until one is set.
      DEFAULT_XML = <<~XML.freeze
        <a:gradFill #{Ns.nsdecls("a")} rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="accent1">
                <a:tint val="100000"/>
                <a:shade val="100000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="accent1">
                <a:tint val="50000"/>
                <a:shade val="100000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin scaled="0"/>
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
    # `a:srcRect`: how much of an image is cropped from each side, as a
    # fraction of its size. Negative values extend past the image edge.
    class CT_RelativeRect < Element
      tag "a:srcRect"
      optional_attr "l", type: SimpleTypes::ST_Percentage, default: 0.0
      optional_attr "t", type: SimpleTypes::ST_Percentage, default: 0.0
      optional_attr "r", type: SimpleTypes::ST_Percentage, default: 0.0
      optional_attr "b", type: SimpleTypes::ST_Percentage, default: 0.0
    end

    class CT_BlipFillProperties < Element
      # The same schema type serves a picture (`p:blipFill`) and a fill.
      tag "a:blipFill", "p:blipFill"
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
