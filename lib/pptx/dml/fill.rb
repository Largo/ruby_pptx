# frozen_string_literal: true

require "pptx/dml/color"
require "pptx/oxml/dml/fill"
require "pptx/enum/dml"

module Pptx
  # The fill of a shape, text run, line or slide background.
  #
  # A fill starts out inherited -- {#type} is nil and nothing is written.
  # Calling {#solid}, {#background} or {#gradient} makes it explicit, after
  # which {#fore_color} and friends become available.
  #
  #   shape.fill.solid
  #   shape.fill.fore_color.rgb = Pptx::RGBColor["C0504D"]
  class FillFormat
    # @param fill_parent [Pptx::Oxml::Element] the element carrying the fill
    #   choice group, e.g. `p:spPr` or `a:rPr`
    def self.from_fill_parent(fill_parent) = new(fill_parent)

    def initialize(fill_parent)
      @parent = fill_parent
    end

    # @return [Pptx::Enum::MSO_FILL_TYPE, nil] nil when the fill is inherited
    def type
      element = fill_element
      return nil if element.nil?

      TYPES[element.nsptag]
    end

    TYPES = {
      "a:noFill" => Enum::MSO_FILL_TYPE::BACKGROUND,
      "a:solidFill" => Enum::MSO_FILL_TYPE::SOLID,
      "a:gradFill" => Enum::MSO_FILL_TYPE::GRADIENT,
      "a:blipFill" => Enum::MSO_FILL_TYPE::PICTURE,
      "a:pattFill" => Enum::MSO_FILL_TYPE::PATTERNED,
      "a:grpFill" => Enum::MSO_FILL_TYPE::GROUP
    }.freeze

    # Make this a solid (flat colour) fill. The colour itself is then set
    # through {#fore_color}.
    def solid
      @parent.get_or_change_to_solidFill
      self
    end

    # Make this fill transparent, so whatever is behind shows through.
    def background
      @parent.get_or_change_to_noFill
      self
    end

    # Make this a gradient fill, with PowerPoint's default stops.
    def gradient
      @parent.get_or_change_to_gradFill
      self
    end

    # Make this a patterned fill.
    def patterned
      @parent.get_or_change_to_pattFill
      self
    end

    # The foreground colour.
    #
    # @raise [Error] unless the fill is solid or patterned
    def fore_color
      case type&.name
      when :SOLID then ColorFormat.from_color_choice_parent(fill_element)
      when :PATTERNED then ColorFormat.from_color_choice_parent(fill_element.get_or_add_fgClr)
      else
        raise Error, "fill type #{type&.name.inspect} has no foreground colour"
      end
    end

    # The background colour of a patterned fill.
    #
    # @raise [Error] unless the fill is patterned
    def back_color
      raise Error, "fill type #{type&.name.inspect} has no background colour" unless type&.name == :PATTERNED

      ColorFormat.from_color_choice_parent(fill_element.get_or_add_bgClr)
    end

    # @return [Pptx::Enum::MSO_PATTERN_TYPE, nil]
    def pattern
      raise Error, "fill type #{type&.name.inspect} is not patterned" unless type&.name == :PATTERNED

      fill_element.prst
    end

    def pattern=(value)
      patterned unless type&.name == :PATTERNED
      fill_element.prst = value
    end

    def inspect = "#<Pptx::FillFormat type=#{type&.name.inspect}>"

    private

    def fill_element = @parent.eg_fillProperties
  end

  # The outline of a shape.
  class LineFormat
    def initialize(parent)
      @parent = parent
    end

    # The `a:ln` element, added if not present.
    def element = @element ||= @parent.get_or_add_ln

    def fill = @fill ||= FillFormat.from_fill_parent(element)

    # Shortcut to the line's solid colour, making the fill solid on first use.
    def color
      fill.solid if fill.type.nil?
      fill.fore_color
    end

    # @return [Pptx::Length, nil] nil when the width is inherited
    def width
      value = element.w
      value.nil? || value.zero? ? nil : value
    end

    def width=(value)
      element.w = value.nil? ? Pptx::Length.emu(0) : value
    end

    def inspect = "#<Pptx::LineFormat width=#{width&.pt}>"
  end
end
