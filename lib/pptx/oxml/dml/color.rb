# frozen_string_literal: true

require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/oxml/simple_types"
require "pptx/enum/dml"

module Pptx
  module Oxml
    # The colour elements share a luminance-adjustment pair: `a:lumMod` scales
    # the colour (a shade) and `a:lumOff` lightens it (a tint).
    module BaseColorElement
      def self.included(base)
        base.class_eval do
          zero_or_one "a:lumMod", successors: []
          zero_or_one "a:lumOff", successors: []
        end
      end

      def add_lumMod(value) = add_lumMod_element(value)

      def add_lumOff(value) = add_lumOff_element(value)

      # Remove any luminance adjustment, returning self.
      def clear_lum
        remove_lumMod
        remove_lumOff
        self
      end

      private

      def add_lumMod_element(value)
        el = get_or_add_lumMod
        el.val = value
        el
      end

      def add_lumOff_element(value)
        el = get_or_add_lumOff
        el.val = value
        el
      end
    end

    # `a:lumMod` and `a:lumOff`.
    class CT_Percentage < Element
      tag "a:lumMod", "a:lumOff"
      required_attr "val", type: SimpleTypes::ST_Percentage
    end

    # `a:srgbClr`, a literal RGB colour.
    class CT_SRgbColor < Element
      include BaseColorElement

      tag "a:srgbClr"
      required_attr "val", type: SimpleTypes::ST_HexColorRGB
    end

    # `a:schemeClr`, a reference to a theme colour.
    class CT_SchemeColor < Element
      include BaseColorElement

      tag "a:schemeClr"
      required_attr "val", type: Enum::MSO_THEME_COLOR
    end

    # `a:hslClr`, a hue/saturation/luminance colour.
    class CT_HslColor < Element
      include BaseColorElement

      tag "a:hslClr"
    end

    # `a:sysClr`, a system colour such as "windowText".
    class CT_SystemColor < Element
      include BaseColorElement

      tag "a:sysClr"
    end

    # `a:prstClr`, a named preset colour.
    class CT_PresetColor < Element
      include BaseColorElement

      tag "a:prstClr"
    end

    # `a:scrgbClr`, a percentage-based RGB colour.
    class CT_ScRgbColor < Element
      include BaseColorElement

      tag "a:scrgbClr"
    end

    # The six-way colour choice shared by everything that carries a colour.
    module ColorChoice
      COLOR_CHOICES = %w[a:scrgbClr a:srgbClr a:hslClr a:sysClr a:schemeClr a:prstClr].freeze

      def self.included(base)
        base.class_eval do
          zero_or_one_choice ColorChoice::COLOR_CHOICES.map { |t| choice(t) },
                             successors: [], as: :eg_colorChoice
        end
      end
    end

    # `a:fgClr` and `a:bgClr`, which are colours and nothing else.
    class CT_Color < Element
      include ColorChoice

      tag "a:fgClr", "a:bgClr"
    end
  end
end
