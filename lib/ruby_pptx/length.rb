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
      def emu(emu)
        new(emu)
      end

      def inches(inches)
        new((inches * EMUS_PER_INCH).to_i)
      end

      def cm(cm)
        new((cm * EMUS_PER_CM).to_i)
      end

      def mm(mm)
        new((mm * EMUS_PER_MM).to_i)
      end

      def pt(points)
        new((points * EMUS_PER_PT).to_i)
      end

      # Hundredths of a point (1/7200 inch); the unit PowerPoint stores font
      # sizes in.
      def centipoints(centipoints)
        new((centipoints * EMUS_PER_CENTIPOINT).to_i)
      end

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

    def inches
      @emu / EMUS_PER_INCH.to_f
    end

    def cm
      @emu / EMUS_PER_CM.to_f
    end

    def mm
      @emu / EMUS_PER_MM.to_f
    end

    def pt
      @emu / EMUS_PER_PT.to_f
    end

    # @return [Integer] whole hundredths of a point, rounded toward negative infinity
    def centipoints
      @emu / EMUS_PER_CENTIPOINT
    end

    def to_i
      @emu
    end
    alias to_int to_i

    def +(other)
      Length.new(@emu + Length.coerce(other).emu)
    end

    def -(other)
      Length.new(@emu - Length.coerce(other).emu)
    end

    def -@
      Length.new(-@emu)
    end

    # Scaling by a plain number; `length * 2` is twice as long.
    def *(factor)
      Length.new((@emu * factor).to_i)
    end

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

    def ==(other)
      (self <=> other)&.zero? || false
    end
    alias eql? ==

    def hash
      @emu.hash
    end

    def zero?
      @emu.zero?
    end

    def to_s
      "#{@emu}emu"
    end

    def inspect
      "#<Pptx::Length #{@emu}emu (#{format("%g", inches)}in)>"
    end
  end

  class << self
    # Convenience constructors, e.g. `Pptx.inches(1.5)`.
    def emu(v)
      Length.emu(v)
    end

    def inches(v)
      Length.inches(v)
    end

    def cm(v)
      Length.cm(v)
    end

    def mm(v)
      Length.mm(v)
    end

    def pt(v)
      Length.pt(v)
    end

    def centipoints(v)
      Length.centipoints(v)
    end
  end
end
