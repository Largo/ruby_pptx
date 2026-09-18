# frozen_string_literal: true

require "ruby_pptx/length"

module Pptx
  # A position on a slide.
  #
  #   Pptx.point(Pptx.inches(1), Pptx.inches(2))
  #   slide.shapes.add_shape(:rectangle, at: origin + offset, size: box)
  #
  # Anywhere this library takes `at:` it has always taken a two-element array,
  # and it still does. A Point defines `to_ary`, so it destructures exactly as
  # that array did -- passing one costs nothing and buys named readers,
  # arithmetic and pattern matching:
  #
  #   case shape.position
  #   in {x:, y:} if x > Pptx.inches(5) then :right_half
  #   end
  #
  # Coordinates are normalised to {Pptx::Length}, so a bare number means EMU,
  # the same as it does in the array form.
  Point = Data.define(:x, :y) do
    # Build from a Point or a two-element array, so callers can accept either.
    def self.from(value)
      value.is_a?(Point) ? value : new(*value)
    end

    def initialize(x:, y:)
      super(x: Length.from(x), y: Length.from(y))
    end

    # Destructuring and splatting, which is what lets a Point stand in for the
    # `[x, y]` array everywhere this library already unpacks one.
    def to_ary = [x, y]

    def +(other)
      other = Point.from(other)
      Point.new(x: x + other.x, y: y + other.y)
    end

    def -(other)
      other = Point.from(other)
      Point.new(x: x - other.x, y: y - other.y)
    end

    def to_s = "(#{x.inches.round(3)}in, #{y.inches.round(3)}in)"
    def inspect = "#<Pptx::Point #{self}>"
  end

  # The extent of a shape: a width and a height.
  #
  # As with {Point}, this stands in for the `[width, height]` array that `size:`
  # has always taken.
  Size = Data.define(:width, :height) do
    def self.from(value)
      value.is_a?(Size) ? value : new(*value)
    end

    def initialize(width:, height:)
      super(width: Length.from(width), height: Length.from(height))
    end

    def to_ary = [width, height]

    # Scale both extents, for `size * 2` or `size * 0.5`.
    def *(factor) = Size.new(width: width * factor, height: height * factor)

    def aspect_ratio = width.emu.to_f / height.emu

    def to_s = "#{width.inches.round(3)}in x #{height.inches.round(3)}in"
    def inspect = "#<Pptx::Size #{self}>"
  end

  class << self
    # @return [Pptx::Point]
    def point(x, y) = Point.new(x: x, y: y)

    # @return [Pptx::Size]
    def size(width, height) = Size.new(width: width, height: height)
  end
end
