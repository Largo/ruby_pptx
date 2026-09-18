# frozen_string_literal: true

require "pptx/text/font_metrics"

module Pptx
  # Finds the largest whole point size at which a string still fits a box.
  #
  # The search is a port of python-pptx's: candidate sizes are tried largest
  # first, and at each size the text is wrapped greedily onto as many lines as
  # it needs and the total height compared against the box. Only the
  # measurement underneath differs -- see {FontMetrics}.
  class TextFitter
    # @param text [String]
    # @param extents [Array(Integer, Integer)] width and height in EMU
    # @param max_size [Integer] largest point size to consider
    # @param font_file [String] path to the font to measure with
    # @return [Integer, nil] nil when the text does not fit at any size
    def self.best_fit_font_size(text, extents:, max_size:, font_file:)
      new(text, extents: extents, font_file: font_file).best_fit(max_size)
    end

    def initialize(text, extents:, font_file:)
      @text = text
      # Extents may arrive as Length; measurements are plain EMU integers.
      @width, @height = extents.map(&:to_i)
      @metrics = font_file.is_a?(FontMetrics) ? font_file : FontMetrics.open(font_file)
    end

    # The largest whole size in 1..max_size that fits, or nil if none does.
    def best_fit(max_size)
      max_size.to_i.downto(1) { |size| return size if fits?(size) }
      nil
    end

    # Wrap +text+ at +point_size+ and report the lines it needs.
    #
    # Exposed because it is the interesting half of the algorithm and is worth
    # being able to check on its own.
    def wrap(point_size)
      lines = []
      remainder = @text.split
      until remainder.empty?
        line, remainder = break_line(remainder, point_size)
        lines << line
      end
      lines
    end

    private

    # The height of a line is measured from "Ty" rather than from the text, so
    # that every line is the same height whether or not it happens to contain
    # an ascender or a descender.
    def fits?(point_size)
      line_height = @metrics.text_extents("Ty", point_size)[1]
      wrap(point_size).size * line_height <= @height
    end

    # Take as many whole words as fit, but never fewer than one -- a single
    # word wider than the box would otherwise loop forever. Such a word makes
    # the line overflow, and the size search then rejects that size on height.
    def break_line(words, point_size)
      taken = 1
      taken += 1 while taken < words.size && fits_width?(words[0, taken + 1], point_size)
      [words[0, taken].join(" "), words[taken..]]
    end

    def fits_width?(words, point_size)
      @metrics.text_extents(words.join(" "), point_size)[0] <= @width
    end
  end
end
