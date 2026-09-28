# frozen_string_literal: true

module Pptx
  # Builds a freeform shape from a sequence of pen movements.
  #
  # Coordinates are given in whatever units suit the drawing -- "local"
  # coordinates -- and +scale+ says how many EMU one local unit is worth. That
  # lets a shape be described in convenient numbers and placed at any size:
  #
  #   shape = slide.shapes.add_freeform(at: [Pptx.inches(1), Pptx.inches(1)],
  #                                     scale: Pptx.inches(1).emu / 100.0) do |f|
  #     f.line_to(100, 0)
  #     f.line_to(50, 100)
  #   end
  #
  # The shape's position and size follow from the path: the bounding box of
  # the points drawn, placed at the origin given to {#convert_to_shape}.
  class FreeformBuilder
    include Enumerable

    # One pen movement. `kind` is :move, :line or :close; a close carries no
    # coordinates.
    Operation = Data.define(:kind, :x, :y)

    def initialize(shapes, start_x, start_y, x_scale, y_scale)
      @shapes = shapes
      @start_x = round(start_x)
      @start_y = round(start_y)
      @x_scale = x_scale
      @y_scale = y_scale
      @operations = []
    end

    def self.new_builder(shapes, start_x, start_y, scale)
      x_scale, y_scale = scale.is_a?(Array) ? scale : [scale, scale]
      new(shapes, start_x, start_y, x_scale, y_scale)
    end

    def each(&)
      @operations.each(&)
    end

    def size
      @operations.size
    end

    # Draw a straight line to (x, y).
    #
    # @return [self] so calls can be chained
    def line_to(x, y)
      @operations << Operation.new(:line, round(x), round(y))
      self
    end

    # Lift the pen and continue from (x, y), starting a new contour.
    def move_to(x, y)
      @operations << Operation.new(:move, round(x), round(y))
      self
    end

    # Join the current contour back to where it started.
    def close
      @operations << Operation.new(:close, nil, nil)
      self
    end

    # Draw a line to each of +vertices+, optionally closing the contour.
    #
    # @param vertices [Array<Array(Numeric, Numeric)>]
    def add_line_segments(vertices, close: true)
      vertices.each { |x, y| line_to(x, y) }
      self.close if close
      self
    end

    # Turn the path into a shape, with the local origin placed at +origin_at+
    # on the slide.
    #
    # May be called more than once to stamp the same geometry in several
    # places.
    #
    # @return [Shape]
    def convert_to_shape(origin_at: [0, 0])
      origin_x, origin_y = origin_at.map { |value| Length.coerce(value).emu }
      sp = @shapes.add_freeform_element(origin_x + left, origin_y + top, width, height)
      path = sp.add_path(local_width, local_height)
      path.move_to(*to_shape_space(@start_x, @start_y))
      @operations.each { |operation| apply(operation, path) }
      @shapes.build_shape(sp)
    end

    # The leftmost extent of the path, in local coordinates.
    #
    # The bounding box need not start at the local origin, so this is what the
    # points are measured from.
    def offset_x
      drawn.map(&:x).push(@start_x).min
    end

    def offset_y
      drawn.map(&:y).push(@start_y).min
    end

    def inspect
      "#<Pptx::FreeformBuilder #{size} operations>"
    end

    private

    # Coordinates may be given as floats for convenience; the format wants
    # integers.
    #
    # Rounded half-to-even, because Python rounds that way and Ruby rounds
    # half-up: a coordinate of 100.5 would otherwise land a unit away from
    # where the reference implementation puts it.
    def round(value)
      value.round(half: :even)
    end

    # Operations that have a position; a close does not.
    def drawn
      @operations.reject { |operation| operation.kind == :close }
    end

    def local_width
      xs = drawn.map(&:x).push(@start_x)
      xs.max - xs.min
    end

    def local_height
      ys = drawn.map(&:y).push(@start_y)
      ys.max - ys.min
    end

    def width
      (local_width * @x_scale).round
    end

    def height
      (local_height * @y_scale).round
    end

    def left
      (offset_x * @x_scale).round
    end

    def top
      (offset_y * @y_scale).round
    end

    # Points inside a path are relative to the shape's top-left corner, not to
    # the local origin.
    def to_shape_space(x, y)
      [x - offset_x, y - offset_y]
    end

    def apply(operation, path)
      case operation.kind
      when :move then path.move_to(*to_shape_space(operation.x, operation.y))
      when :line then path.line_to(*to_shape_space(operation.x, operation.y))
      when :close then path.close_contour
      end
    end
  end
end
