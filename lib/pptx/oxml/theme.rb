# frozen_string_literal: true

require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/oxml/simple_types"
require "pptx/oxml/dml/color"
require "pptx/oxml/text"

module Pptx
  module Oxml
    # `a:theme`, the root of a theme part.
    #
    # Only the parts a caller can reasonably edit are modelled: the colour
    # scheme and the font scheme. The format scheme -- fills, lines, effects
    # and background fills, keyed by the `idx` values that shapes and slide
    # backgrounds refer to -- is carried through from the base theme unchanged.
    # See PORTING.md for why.
    class CT_OfficeStyleSheet < Element
      tag "a:theme"
      one_and_only_one "a:themeElements"
      required_attr "name", type: SimpleTypes::XsdString

      def self.new_default = Oxml.parse_from_template("theme")

      def clrScheme = themeElements.clrScheme
      def fontScheme = themeElements.fontScheme
    end

    # `a:themeElements`, the four schemes that make up a theme.
    class CT_BaseStyles < Element
      tag "a:themeElements"
      one_and_only_one "a:clrScheme"
      one_and_only_one "a:fontScheme"
      one_and_only_one "a:fmtScheme"
    end

    # `a:clrScheme`, the twelve theme colours.
    class CT_ColorScheme < Element
      tag "a:clrScheme"
      required_attr "name", type: SimpleTypes::XsdString

      # In schema order, which is also the order PowerPoint shows them in.
      SLOTS = %w[dk1 lt1 dk2 lt2 accent1 accent2 accent3 accent4 accent5
                 accent6 hlink folHlink].freeze

      SLOTS.each do |slot|
        one_and_only_one "a:#{slot}"
      end

      def slot(name)
        raise NotFoundError, "no theme colour named #{name}" unless SLOTS.include?(name.to_s)

        send(name.to_s)
      end
    end

    # One slot of a colour scheme, such as `a:accent1`.
    #
    # Every slot holds the same six-way colour choice as any other colour in
    # DrawingML, so {Pptx::ColorFormat} drives them unchanged.
    class CT_ThemeColor < Element
      include ColorChoice
      tag(*CT_ColorScheme::SLOTS.map { |slot| "a:#{slot}" })
    end

    # `a:fontScheme`, the major (heading) and minor (body) font collections.
    class CT_FontScheme < Element
      tag "a:fontScheme"
      required_attr "name", type: SimpleTypes::XsdString
      one_and_only_one "a:majorFont"
      one_and_only_one "a:minorFont"
    end

    # `a:majorFont` and `a:minorFont`.
    #
    # The Latin typeface is the one that matters here; the East Asian and
    # complex-script slots are required by the schema and are left as the base
    # theme has them unless set.
    class CT_FontCollection < Element
      tag "a:majorFont", "a:minorFont"
      one_and_only_one "a:latin"
      one_and_only_one "a:ea"
      one_and_only_one "a:cs"

      def typeface = latin.typeface

      def typeface=(value)
        latin.typeface = value
        value
      end
    end
  end
end
