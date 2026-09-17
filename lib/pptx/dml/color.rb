# frozen_string_literal: true

require "pptx/element_proxy"
require "pptx/oxml/dml/color"
require "pptx/enum/dml"

module Pptx
  # An RGB colour.
  #
  #   Pptx::RGBColor.new(60, 47, 128)
  #   Pptx::RGBColor["3C2F80"]
  class RGBColor
    attr_reader :r, :g, :b

    # Parse a hex string such as "3C2F80".
    def self.from_string(hex)
      unless /\A\h{6}\z/.match?(hex.to_s)
        raise ArgumentError, "expected a six-digit hex string, got #{hex.inspect}"
      end

      new(hex[0, 2].to_i(16), hex[2, 2].to_i(16), hex[4, 2].to_i(16))
    end

    class << self
      alias [] from_string
    end

    def initialize(r, g, b)
      [r, g, b].each do |component|
        unless component.is_a?(Integer) && component.between?(0, 255)
          raise ArgumentError, "RGBColor takes three integers 0-255, got #{[r, g, b].inspect}"
        end
      end
      @r = r
      @g = g
      @b = b
      freeze
    end

    def to_a = [@r, @g, @b]

    # The uppercase hex form OOXML uses, e.g. "3C2F80".
    def to_s = format("%02X%02X%02X", @r, @g, @b)

    def ==(other) = other.is_a?(RGBColor) && other.to_a == to_a
    alias eql? ==

    def hash = to_a.hash

    def inspect = "#<Pptx::RGBColor #{self}>"
  end

  # The colour of a font, fill or line.
  #
  # A colour may be an explicit RGB value, a reference to a theme colour, or
  # absent -- in which case it is inherited and {#type} is nil.
  class ColorFormat
    # @param color_choice_parent [Pptx::Oxml::Element] the element holding the
    #   colour choice, e.g. `a:solidFill`
    def self.from_color_choice_parent(color_choice_parent) = new(color_choice_parent)

    def initialize(color_choice_parent)
      @parent = color_choice_parent
    end

    # @return [Pptx::Enum::MSO_COLOR_TYPE, nil] nil when no colour is set here
    def type
      element = color_element
      return nil if element.nil?

      case element.nsptag
      when "a:srgbClr" then Enum::MSO_COLOR_TYPE::RGB
      when "a:schemeClr" then Enum::MSO_COLOR_TYPE::SCHEME
      when "a:hslClr" then Enum::MSO_COLOR_TYPE::HSL
      when "a:prstClr" then Enum::MSO_COLOR_TYPE::PRESET
      when "a:scrgbClr" then Enum::MSO_COLOR_TYPE::SCRGB
      when "a:sysClr" then Enum::MSO_COLOR_TYPE::SYSTEM
      end
    end

    # @return [RGBColor, nil] nil unless an explicit RGB colour is set
    def rgb
      element = color_element
      return nil unless element&.nsptag == "a:srgbClr"

      RGBColor.from_string(element.val)
    end

    def rgb=(value)
      value = RGBColor.from_string(value) if value.is_a?(String)
      raise TypeError, "expected an RGBColor, got #{value.class}" unless value.is_a?(RGBColor)

      @parent.get_or_change_to_srgbClr.val = value.to_s
      value
    end

    # @return [Pptx::Enum::MSO_THEME_COLOR, nil] nil unless a theme colour is set
    def theme_color
      element = color_element
      return nil unless element&.nsptag == "a:schemeClr"

      element.val
    end

    def theme_color=(value)
      @parent.get_or_change_to_schemeClr.val = Enum::MSO_THEME_COLOR.fetch(value)
      value
    end

    # A luminance adjustment between -1.0 and 1.0: negative is darker (a
    # shade), positive is lighter (a tint).
    def brightness
      element = color_element
      return 0.0 if element.nil?

      lum_off = element.lumOff
      return lum_off.val if lum_off

      lum_mod = element.lumMod
      return lum_mod.val - 1.0 if lum_mod

      0.0
    end

    def brightness=(value)
      unless value.is_a?(Numeric) && value.between?(-1.0, 1.0)
        raise ArgumentError, "brightness must be a number between -1.0 and 1.0, got #{value.inspect}"
      end

      element = color_element
      if element.nil?
        raise Error, "cannot set brightness before a colour; set rgb or theme_color first"
      end

      element.clear_lum
      if value.positive?
        element.add_lumMod(1.0 - value)
        element.add_lumOff(value)
      elsif value.negative?
        element.add_lumMod(1.0 + value)
      end
      value
    end

    def inspect = "#<Pptx::ColorFormat type=#{type&.name.inspect} rgb=#{rgb&.to_s.inspect}>"

    private

    def color_element = @parent.eg_colorChoice
  end
end
