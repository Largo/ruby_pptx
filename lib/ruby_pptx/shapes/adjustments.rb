# frozen_string_literal: true

require "ruby_pptx/autoshape_spec"

module Pptx
  # The adjustment handles of an auto shape -- the yellow diamonds that set
  # how round a rounded rectangle is, or how deep a chevron's notch goes.
  #
  #   shape.adjustments[0]          #=> 0.16667
  #   shape.adjustments[0] = 0.4
  #
  # Values are fractions: 0.5 is half way. Many shapes allow values below
  # 0.0 or above 1.0 at extreme proportions, so neither bound is enforced.
  #
  # Every handle the shape type defines is always present, at its default
  # until set. Setting any one writes all of them, as PowerPoint does.
  class Adjustments
    include Enumerable

    # The file stores adjustments on a scale of 100,000.
    SCALE = 100_000.0

    Adjustment = Struct.new(:name, :default, :actual) do
      def raw = actual || default
    end

    # @param preset_geometry [Pptx::Oxml::Element, nil] the shape's
    #   `a:prstGeom`; nil for a freeform, which has no handles
    def initialize(preset_geometry)
      @geometry = preset_geometry
      @adjustments = initial_adjustments
    end

    # @return [Float, nil] nil past the last handle, as Array does
    def [](index)
      adjustment = @adjustments[index]
      adjustment && (adjustment.raw / SCALE)
    end

    def []=(index, value)
      unless value.is_a?(Numeric)
        raise ArgumentError,
              "adjustment values must be numeric, got #{value.inspect}"
      end

      adjustment = @adjustments[index] or
        raise IndexError, "this shape has #{size} adjustment#{"s" unless size == 1}, not #{index + 1}"
      # Truncated, not rounded, to match python-pptx: 0.29 is stored as 28999.
      adjustment.actual = (value * SCALE).to_i
      @geometry.rewrite_guides(@adjustments.map { |a| [a.name, a.raw] })
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @adjustments.each { |a| yield a.raw / SCALE }
      self
    end

    def size = @adjustments.size
    alias length size

    def inspect = "#<Pptx::Adjustments #{to_a.inspect}>"

    private

    # The type's defaults, overlaid with any values the file records. A
    # guide naming no handle of this type is ignored rather than trusted.
    def initial_adjustments
      return [] if @geometry.nil?

      adjustments = AutoShapeSpec.default_adjustments(@geometry.prst.name).map do |name, default|
        Adjustment.new(name, default)
      end
      by_name = adjustments.to_h { |a| [a.name, a] }
      @geometry.gd_list.each do |gd|
        target = by_name[gd.name] or next
        target.actual = Integer(gd.fmla.delete_prefix("val "), 10)
      end
      adjustments
    end
  end
end
