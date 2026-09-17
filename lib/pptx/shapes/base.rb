# frozen_string_literal: true

require "pptx/element_proxy"
require "pptx/enum/shapes"

module Pptx
  # A shape on a slide, layout or master.
  #
  # Subclasses cover the specific kinds: {Shape} for auto shapes and text
  # boxes, {Picture}, {GraphicFrame}, {Connector} and {GroupShape}.
  class BaseShape
    attr_reader :element, :parent

    def initialize(shape_element, parent)
      @element = shape_element
      @parent = parent
    end

    def ==(other) = other.is_a?(BaseShape) && other.element == @element
    alias eql? ==

    def hash = @element.hash

    # The package part containing this shape.
    def part = @parent.part

    # The drawing-object id, unique within the slide.
    def shape_id = @element.shape_id

    def name = @element.shape_name

    def name=(value)
      @element.nvXxPr.cNvPr.name = value.to_s
      value
    end

    # @return [Length, nil] nil when the shape inherits its position
    def left = @element.x

    def left=(value)
      @element.x = value
      value
    end

    def top = @element.y

    def top=(value)
      @element.y = value
      value
    end

    def width = @element.cx

    def width=(value)
      @element.cx = value
      value
    end

    def height = @element.cy

    def height=(value)
      @element.cy = value
      value
    end

    # Clockwise rotation in degrees.
    def rotation = @element.rot

    def rotation=(value)
      @element.rot = value
      value
    end

    def placeholder? = @element.placeholder?

    # Placeholder position and type, or nil when this is not a placeholder.
    def placeholder_format
      placeholder? ? PlaceholderFormat.new(@element.ph) : nil
    end

    # Overridden by the subclasses that can actually hold these.
    def text_frame? = false
    def chart? = false
    def table? = false

    # @return [Pptx::Enum::MSO_SHAPE_TYPE]
    def shape_type = raise(NotImplementedError, "#{self.class} must implement #shape_type")

    def inspect = "#<#{self.class.name} id=#{shape_id} #{name.inspect}>"
  end

  # The placeholder-specific properties of a shape: which placeholder it is
  # and what kind.
  class PlaceholderFormat < ElementProxy
    # The `idx` that ties a slide placeholder to the layout one it inherits
    # from. The title placeholder is always 0.
    def idx = @element.idx

    # @return [Pptx::Enum::PP_PLACEHOLDER]
    def type = @element.type

    def orientation = @element.orient

    def size = @element.sz

    def inspect = "#<Pptx::PlaceholderFormat idx=#{idx} type=#{type}>"
  end
end
