# frozen_string_literal: true

module Pptx
  # Text measurements read straight out of a TrueType or OpenType font file.
  #
  # This is what {TextFitter} measures with. python-pptx asks Pillow to
  # rasterize the text and reports the ink bounding box of the result; there is
  # no Ruby equivalent of that, so this reads the glyph outlines' own bounding
  # boxes from the font instead. See PORTING.md for how far the two diverge --
  # in short, well under a point size.
  #
  # Only the tables needed for measurement are parsed. Composite glyphs are not
  # recursed into: their bounding box is stored on the glyph itself, which is
  # all this needs.
  class FontMetrics
    EMU_PER_INCH = 914_400
    POINTS_PER_INCH = 72.0

    class << self
      # Fonts are immutable and parsing one is not free, so they are cached.
      def open(path)
        real = File.realpath(path)
        cache[real] ||= new(real)
      end

      def clear_cache = cache.clear

      private

      def cache = @cache ||= {}
    end

    # @param path [String] a .ttf or .otf file
    def initialize(path)
      @path = path
      @data = File.binread(path).b
      @tables = read_table_directory
      require_tables!
      @units_per_em = u16(@tables["head"] + 18)
      @index_to_loc_format = u16(@tables["head"] + 50)
      @num_glyphs = u16(@tables["maxp"] + 4)
      @num_h_metrics = u16(@tables["hhea"] + 34)
      @cmap = read_cmap
      @glyph_boxes = {}
    end

    attr_reader :path, :units_per_em

    # The size of +text+ rendered at +point_size+, as [width, height] in EMU.
    #
    # The height is the ink height, so it depends on which glyphs are present:
    # "Ty" is taller than "ace". That is deliberate -- it is how python-pptx
    # measures, and {TextFitter} relies on it by always measuring "Ty".
    def text_extents(text, point_size)
      box = ink_box(text)
      return [0, 0] if box.nil?

      # Width is measured from the pen's starting position rather than from the
      # first mark, which is what Pillow reports and so what python-pptx fits
      # against: it counts a leading space, and the first glyph's left side
      # bearing. Only ink that starts left of the origin pushes the edge out.
      # Height is pure ink, top mark to bottom mark.
      left = [box[0], 0].min
      [to_emu(box[2] - left, point_size), to_emu(box[3] - box[1], point_size)]
    end

    def inspect = "#<Pptx::FontMetrics #{File.basename(@path)} upem=#{@units_per_em}>"

    private

    # Laying the glyphs out along the baseline and unioning their boxes.
    # Kerning is not applied; see PORTING.md.
    def ink_box(text)
      pen = 0
      left = bottom = right = top = nil
      text.each_char do |char|
        gid = glyph_id(char.ord)
        box = glyph_box(gid)
        if box
          left = min(left, pen + box[0])
          bottom = min(bottom, box[1])
          right = max(right, pen + box[2])
          top = max(top, box[3])
        end
        pen += advance(gid)
      end
      left && [left, bottom, right, top]
    end

    def min(a, b) = a.nil? || b < a ? b : a
    def max(a, b) = a.nil? || b > a ? b : a

    # Font units are relative to the em square, so scaling by the point size
    # gives points, which are 1/72 inch by definition.
    def to_emu(units, point_size)
      (units * point_size / @units_per_em.to_f / POINTS_PER_INCH * EMU_PER_INCH).to_i
    end

    def glyph_id(codepoint) = @cmap[codepoint] || 0

    # A font may store one advance for a run of trailing glyphs, in which case
    # hmtx is truncated and the last entry stands for all of them.
    def advance(gid)
      u16(@tables["hmtx"] + ([gid, @num_h_metrics - 1].min * 4))
    end

    # @return [Array(Integer, Integer, Integer, Integer), nil] xMin, yMin,
    #   xMax, yMax in font units, or nil for a glyph with no outline
    def glyph_box(gid)
      return @glyph_boxes[gid] if @glyph_boxes.key?(gid)

      @glyph_boxes[gid] = compute_glyph_box(gid)
    end

    def compute_glyph_box(gid)
      return cff_glyph_box(gid) if @tables["glyf"].nil?
      return nil if gid >= @num_glyphs

      from = loca(gid)
      # A glyph with no contours, such as the space, has a zero-length entry.
      return nil if from == loca(gid + 1)

      at = @tables["glyf"] + from
      [s16(at + 2), s16(at + 4), s16(at + 6), s16(at + 8)]
    end

    # OpenType/CFF fonts have no glyf table, and reading their charstrings to
    # find outlines is a disproportionate amount of machinery for a measurement
    # that only has to be close. Fall back to the advance width and the font's
    # overall vertical extents.
    def cff_glyph_box(gid)
      width = advance(gid)
      return nil if width.zero?

      ascender = s16(@tables["hhea"] + 4)
      descender = s16(@tables["hhea"] + 6)
      [0, descender, width, ascender]
    end

    def loca(index)
      base = @tables["loca"]
      @index_to_loc_format.zero? ? u16(base + (index * 2)) * 2 : u32(base + (index * 4))
    end

    def u16(at) = @data.byteslice(at, 2).unpack1("n")
    def s16(at) = @data.byteslice(at, 2).unpack1("s>")
    def u32(at) = @data.byteslice(at, 4).unpack1("N")

    def read_table_directory
      raise Error, "#{@path} is not a TrueType or OpenType font" unless sfnt_version?

      u16(4).times.to_h do |i|
        entry = @data.byteslice(12 + (i * 16), 16)
        [entry.byteslice(0, 4), entry.byteslice(8, 4).unpack1("N")]
      end
    end

    # A TrueType outline font, an OpenType/CFF font, or an Apple-flavoured one.
    def sfnt_version?
      %W[\x00\x01\x00\x00 OTTO true].include?(@data.byteslice(0, 4))
    end

    def require_tables!
      missing = %w[head maxp hhea hmtx cmap] - @tables.keys
      return if missing.empty?

      raise Error, "#{@path} is missing required font tables: #{missing.join(', ')}"
    end

    # A Unicode BMP subtable is enough: presentation text outside the BMP falls
    # back to glyph 0, which measures as .notdef rather than crashing.
    def read_cmap
      base = @tables["cmap"]
      offset = nil
      u16(base + 2).times do |i|
        platform, encoding, sub = @data.byteslice(base + 4 + (i * 8), 8).unpack("nnN")
        next unless unicode_bmp?(platform, encoding)
        next unless u16(base + sub) == 4

        offset = base + sub
        break
      end
      offset.nil? ? {} : read_cmap_format4(offset)
    end

    def unicode_bmp?(platform, encoding)
      (platform == 3 && encoding == 1) || (platform.zero? && encoding <= 4)
    end

    def read_cmap_format4(offset)
      seg_x2 = u16(offset + 6)
      ends = @data.byteslice(offset + 14, seg_x2).unpack("n*")
      # A reserved padding word sits between the end and start arrays.
      starts = @data.byteslice(offset + 16 + seg_x2, seg_x2).unpack("n*")
      deltas = @data.byteslice(offset + 16 + (seg_x2 * 2), seg_x2).unpack("n*")
      ranges_at = offset + 16 + (seg_x2 * 3)
      ranges = @data.byteslice(ranges_at, seg_x2).unpack("n*")

      map = {}
      (seg_x2 / 2).times do |i|
        (starts[i]..ends[i]).each do |code|
          # The final segment is a sentinel ending at 0xFFFF.
          next if code == 0xFFFF

          gid = segment_glyph(i, code, starts, deltas, ranges, ranges_at)
          map[code] = gid if gid
        end
      end
      map
    end

    # idRangeOffset is a byte offset from its own slot in the array, which is
    # why the segment index has to be folded back in.
    def segment_glyph(index, code, starts, deltas, ranges, ranges_at)
      return (code + deltas[index]) % 65_536 if ranges[index].zero?

      gid = u16(ranges_at + (index * 2) + ranges[index] + ((code - starts[index]) * 2))
      gid.zero? ? nil : (gid + deltas[index]) % 65_536
    end
  end
end
