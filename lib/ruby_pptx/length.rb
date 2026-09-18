# frozen_string_literal: true

require "ruby_pptx/pattern_matching"

module Pptx
  # A distance, stored canonically in English Metric Units (EMU).
  #
  # OOXML expresses nearly every distance in EMU, an integer unit chosen so that
  # inches, centimetres and points all divide it evenly. `Length` keeps that
  # integer as the single source of truth and converts on the way out.
  #
  #   Pptx::Length.inches(1).emu  #=> 914400
  #   Pptx::Length.cm(2.54).pt    #=> 72.0
  #
  # Instances are immutable, `Comparable`, and compare against bare Integers as
  # EMU, so `Pptx::Length.inches(1) == 914_400` holds.
  class Length
    include Comparable
    include PatternMatching

    # `in {inches:}` converts on demand, so a pattern can ask in whatever unit
    # reads best at the point of use.
    pattern_keys :emu, :inches, :cm, :mm, :pt, :centipoints

    EMUS_PER_INCH       = 914_400
    EMUS_PER_CENTIPOINT = 127
    EMUS_PER_CM         = 360_000
    EMUS_PER_MM         = 36_000
    EMUS_PER_PT         = 12_700

    class << self
      # @param emu [Integer] distance in English Metric Units
      # Accept either a Length or a bare number of EMU, so a caller may hand
      # over whichever it has without checking first.
      def from(value) = value.is_a?(Length) ? value : emu(value)

      def emu(emu) = new(emu)

      def inches(inches) = new((inches * EMUS_PER_INCH).to_i)
      def cm(cm)         = new((cm * EMUS_PER_CM).to_i)
      def mm(mm)         = new((mm * EMUS_PER_MM).to_i)
      def pt(points)     = new((points * EMUS_PER_PT).to_i)

      # Hundredths of a point (1/7200 inch); the unit PowerPoint stores font
      # sizes in.
      def centipoints(centipoints) = new((centipoints * EMUS_PER_CENTIPOINT).to_i)

      # Coerce +value+ to a Length, treating a bare Integer as EMU.
      #
      # @return [Length, nil] nil when +value+ is nil
      def coerce(value)
        case value
        when nil     then nil
        when Length  then value
        when Integer then new(value)
        else
          raise TypeError, "expected Pptx::Length or Integer (EMU), got #{value.class}"
        end
      end
    end

    # @return [Integer] the distance in English Metric Units
    attr_reader :emu

    def initialize(emu)
      @emu = Integer(emu)
      freeze
    end

    def inches = @emu / EMUS_PER_INCH.to_f
    def cm     = @emu / EMUS_PER_CM.to_f
    def mm     = @emu / EMUS_PER_MM.to_f
    def pt     = @emu / EMUS_PER_PT.to_f

    # @return [Integer] whole hundredths of a point, rounded toward negative infinity
    def centipoints = @emu / EMUS_PER_CENTIPOINT

    def to_i = @emu
    alias to_int to_i

    def +(other) = Length.new(@emu + Length.coerce(other).emu)
    def -(other) = Length.new(@emu - Length.coerce(other).emu)
    def -@ = Length.new(-@emu)

    # Scaling by a plain number; `length * 2` is twice as long.
    def *(factor) = Length.new((@emu * factor).to_i)

    # Deliberately no #coerce, so `2 * length` raises rather than working.
    #
    # Ruby's coerce protocol cannot see which operator is being applied, so the
    # pair it returns has to be right for all of them. Returning [self, other]
    # would make `2 * length` correct but silently turn `2 - length` into
    # `length - 2` -- the wrong sign. A TypeError the caller can fix by writing
    # `length * 2` beats an answer that is quietly negative.

    def <=>(other)
      other = Length.coerce(other) if other.is_a?(Integer)
      return nil unless other.is_a?(Length)

      @emu <=> other.emu
    end

    def ==(other) = (self <=> other)&.zero? || false
    alias eql? ==

    def hash = @emu.hash

    def zero? = @emu.zero?

    def to_s = "#{@emu}emu"

    def inspect = "#<Pptx::Length #{@emu}emu (#{format("%g", inches)}in)>"
  end

  class << self
    # Convenience constructors, e.g. `Pptx.inches(1.5)`.
    def emu(v)         = Length.emu(v)
    def inches(v)      = Length.inches(v)
    def cm(v)          = Length.cm(v)
    def mm(v)          = Length.mm(v)
    def pt(v)          = Length.pt(v)
    def centipoints(v) = Length.centipoints(v)
  end
end
