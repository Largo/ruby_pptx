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
      DEFAULT_XML = <<~XML
        <a:gradFill #{Ns.nsdecls('a')} rotWithShape="1">
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
