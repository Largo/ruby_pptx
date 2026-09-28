# frozen_string_literal: true

require "ruby_pptx/dml/color"
require "ruby_pptx/oxml/dml/fill"
require "ruby_pptx/enum/dml"
require "ruby_pptx/element_proxy"
require "ruby_pptx/sliceable"

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
    def self.from_fill_parent(fill_parent)
      new(fill_parent)
    end

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

    # The angle of a linear gradient in degrees, counter-clockwise from
    # pointing right -- the way PowerPoint's dialog and trigonometry both
    # count, although the file stores it clockwise. nil when inherited.
    #
    # @raise [Error] unless this is a gradient fill, or if it is not linear
    def gradient_angle
      gradient = require_gradient
      raise Error, "not a linear gradient" unless gradient.path.nil?

      clockwise = gradient.lin&.ang
      return nil if clockwise.nil?

      clockwise.zero? ? 0.0 : 360.0 - clockwise
    end

    def gradient_angle=(value)
      linear = require_gradient.lin
      raise Error, "not a linear gradient" if linear.nil?

      linear.ang = 360.0 - value
    end

    # The colour stops the gradient passes through, in order.
    #
    # @return [GradientStops]
    def gradient_stops
      GradientStops.new(require_gradient.get_or_add_gsLst)
    end

    def inspect
      "#<Pptx::FillFormat type=#{type&.name.inspect}>"
    end

    private

    def fill_element
      @parent.eg_fillProperties
    end

    def require_gradient
      raise Error, "fill type #{type&.name.inspect} is not a gradient" unless type&.name == :GRADIENT

      fill_element
    end
  end

  # The outline of a shape.
  # The stops of a gradient fill.
  #
  #   stops = shape.fill.gradient_stops
  #   stops[0].color.rgb = Pptx::RGBColor["1F497D"]
  #   stops[1].position = 0.75
  class GradientStops
    include Enumerable
    include Sliceable

    def initialize(gs_lst)
      @element = gs_lst
    end

    def each(&)
      return enum_for(:each) { size } unless block_given?

      @element.gs_list.each { |gs| yield GradientStop.new(gs) }
      self
    end

    def size
      @element.gs_list.size
    end
    alias length size

    def [](index, length = nil)
      slice_members(@element.gs_list, index, length) { |gs| GradientStop.new(gs) }
    end

    def inspect
      "#<Pptx::GradientStops size=#{size}>"
    end
  end

  # One colour stop in a gradient.
  class GradientStop < ElementProxy
    def color
      @color ||= ColorFormat.from_color_choice_parent(@element)
    end

    # Where along the gradient this stop sits, from 0.0 at the start to 1.0
    # at the end.
    def position
      @element.pos
    end

    def position=(value)
      @element.pos = Float(value)
    end

    def inspect
      "#<Pptx::GradientStop position=#{position}>"
    end
  end

  class LineFormat
    def initialize(parent)
      @parent = parent
    end

    # The `a:ln` element, added if not present.
    def element
      @element ||= @parent.get_or_add_ln
    end

    def fill
      @fill ||= FillFormat.from_fill_parent(element)
    end

    # Shortcut to the line's solid colour, making the fill solid on first use.
    def color
      fill.solid if fill.type.nil?
      fill.fore_color
    end

    # @return [Pptx::Length, nil] nil when the width is inherited
    #
    # Reading never adds an `a:ln`; only assigning does.
    def width
      value = @parent.ln&.w
      value.nil? || value.zero? ? nil : value
    end

    # The dash pattern, or nil when it is inherited.
    #
    # @return [Pptx::Enum::Member, nil] a member of MSO_LINE_DASH_STYLE
    def dash_style
      @parent.ln&.prstDash_val
    end

    # Set the dash pattern, or restore inheritance with nil -- which also
    # drops a custom dash, since either one would override the inherited
    # style.
    def dash_style=(value)
      if value.nil?
        line = @parent.ln or return
        line.remove_prstDash
        line.remove_custDash
      else
        element.prstDash_val = Enum::MSO_LINE_DASH_STYLE.fetch(value)
      end
    end

    def width=(value)
      element.w = value.nil? ? Pptx::Length.emu(0) : value
    end

    def inspect
      "#<Pptx::LineFormat width=#{width&.pt}>"
    end
  end
end
