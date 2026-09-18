# frozen_string_literal: true

require "digest"
require "stringio"
require "ruby_pptx/errors"
require "ruby_pptx/length"
require "ruby_pptx/opc/constants"

module Pptx
  # An image, as a value object over its bytes.
  #
  # Format, pixel size and resolution are read straight from the file header.
  # python-pptx uses Pillow for this; a handful of header formats is a much
  # smaller thing to carry than an image library, and the specs check the
  # results against Pillow for every supported format.
  class Image
    # Canonical extension per format. The extension of the file the image came
    # from is ignored: what matters is what the bytes actually are.
    EXT_FOR_FORMAT = {
      BMP: "bmp", GIF: "gif", JPEG: "jpg", PNG: "png", TIFF: "tiff", WMF: "wmf",
      SVG: "svg"
    }.freeze

    # Not in the OPC spec's table; Office writes this for an SVG part.
    CONTENT_TYPE_SVG = "image/svg+xml"

    CONTENT_TYPE_FOR_EXT = {
      "bmp" => Opc::CONTENT_TYPE::BMP,
      "emf" => Opc::CONTENT_TYPE::X_EMF,
      "gif" => Opc::CONTENT_TYPE::GIF,
      "jpe" => Opc::CONTENT_TYPE::JPEG,
      "jpeg" => Opc::CONTENT_TYPE::JPEG,
      "jpg" => Opc::CONTENT_TYPE::JPEG,
      "png" => Opc::CONTENT_TYPE::PNG,
      "tif" => Opc::CONTENT_TYPE::TIFF,
      "tiff" => Opc::CONTENT_TYPE::TIFF,
      "wdp" => Opc::CONTENT_TYPE::MS_PHOTO,
      "wmf" => Opc::CONTENT_TYPE::X_WMF,
      "svg" => CONTENT_TYPE_SVG
    }.freeze

    # Resolution assumed when the file does not say.
    DEFAULT_DPI = 72

    # A dpi outside this range is not believable and is replaced by the default.
    DPI_RANGE = (1..2048)

    EMU_PER_INCH = 914_400

    attr_reader :blob, :filename

    class << self
      def from_blob(blob, filename = nil) = new(blob, filename)

      # Load from a path or an IO stream.
      def from_file(image_file)
        if image_file.is_a?(String)
          new(File.binread(image_file), File.basename(image_file))
        else
          image_file.rewind if image_file.respond_to?(:rewind)
          new(image_file.read, nil)
        end
      end
    end

    def initialize(blob, filename = nil)
      @blob = blob
      @filename = filename
    end

    # @return [Symbol] :PNG, :JPEG, :GIF, :BMP, :TIFF or :WMF
    def format = header.fetch(:format)

    # True for a vector image, which has no pixel size of its own and needs a
    # raster fallback to be placed in a slide.
    def vector? = format == :SVG

    # @return [Array(Integer, Integer)] width and height in pixels
    def size = [header.fetch(:width), header.fetch(:height)]

    def width_px = header.fetch(:width)

    def height_px = header.fetch(:height)

    # @return [Array(Integer, Integer)] horizontal and vertical dots per inch
    def dpi
      [normalize_dpi(header[:horz_dpi]), normalize_dpi(header[:vert_dpi])]
    end

    # The canonical extension for this image's actual format.
    def ext
      EXT_FOR_FORMAT.fetch(format) do
        raise Error, "unsupported image format #{format}"
      end
    end

    def content_type = CONTENT_TYPE_FOR_EXT.fetch(ext)

    def sha1 = @sha1 ||= Digest::SHA1.hexdigest(@blob)

    # Native size in EMU, derived from the pixel size and resolution.
    #
    # @return [Array(Pptx::Length, Pptx::Length)]
    def native_size
      raise Error, "an SVG has no native pixel size; size it from its raster fallback" if vector?

      horz_dpi, vert_dpi = dpi
      [Length.emu((EMU_PER_INCH * width_px / horz_dpi).to_i),
       Length.emu((EMU_PER_INCH * height_px / vert_dpi).to_i)]
    end

    def inspect = "#<Pptx::Image #{format} #{width_px}x#{height_px} #{dpi.join("x")}dpi>"

    private

    def normalize_dpi(value)
      return DEFAULT_DPI if value.nil?

      rounded = Float(value).round
      DPI_RANGE.cover?(rounded) ? rounded : DEFAULT_DPI
    rescue ArgumentError, TypeError
      DEFAULT_DPI
    end

    def header
      @header ||= ImageHeader.parse(@blob)
    end
  end

  # Reads format, pixel dimensions and resolution out of an image file header.
  module ImageHeader
    PNG_SIGNATURE = "\x89PNG\r\n\x1a\n".b
    # A pHYs chunk gives pixels per unit; unit 1 is the metre.
    METRES_PER_INCH = 0.0254

    module_function

    # @return [Hash] {format:, width:, height:, horz_dpi:, vert_dpi:}
    def parse(blob)
      blob = blob.b
      return parse_png(blob) if blob.start_with?(PNG_SIGNATURE)
      return parse_jpeg(blob) if blob.start_with?("\xFF\xD8".b)
      return parse_gif(blob) if blob.start_with?("GIF87a".b, "GIF89a".b)
      return parse_bmp(blob) if blob.start_with?("BM".b)
      return parse_tiff(blob) if blob.start_with?("II*\x00".b, "MM\x00*".b)
      return { format: :WMF, width: 0, height: 0 } if blob.start_with?("\xD7\xCD\xC6\x9A".b)
      return { format: :SVG } if svg?(blob)

      raise Error, "unrecognized image format"
    end

    # An SVG is XML, so it is recognized by its root element rather than by a
    # magic number. Only the first kilobyte is searched, which covers any
    # plausible declaration, comment or doctype preamble.
    def svg?(blob)
      head = blob.byteslice(0, 1024)
      head.include?("<svg".b) || (head.include?("<?xml".b) && blob.include?("<svg".b))
    end

    # PNG is a signature followed by length-prefixed chunks. IHDR carries the
    # dimensions and is always first; pHYs, if present, carries resolution.
    def parse_png(blob)
      result = { format: :PNG }
      offset = PNG_SIGNATURE.bytesize

      while offset + 8 <= blob.bytesize
        length = blob.byteslice(offset, 4).unpack1("N")
        type = blob.byteslice(offset + 4, 4)
        data = blob.byteslice(offset + 8, length)
        break if data.nil?

        case type
        when "IHDR"
          result[:width], result[:height] = data.unpack("NN")
        when "pHYs"
          ppu_x, ppu_y, unit = data.unpack("NNC")
          if unit == 1
            result[:horz_dpi] = ppu_x * METRES_PER_INCH
            result[:vert_dpi] = ppu_y * METRES_PER_INCH
          end
        when "IEND"
          break
        end
        offset += 12 + length
      end
      result
    end

    # JPEG is a stream of marker segments. SOFn carries the dimensions; the
    # JFIF APP0 segment, when present, carries the resolution.
    def parse_jpeg(blob)
      result = { format: :JPEG }
      offset = 2

      while offset + 4 <= blob.bytesize
        # Markers may be preceded by any number of 0xFF fill bytes.
        offset += 1 while offset < blob.bytesize && blob.getbyte(offset) == 0xFF
        marker = blob.getbyte(offset)
        offset += 1
        break if marker.nil?
        # Standalone markers carry no payload.
        next if (0xD0..0xD9).cover?(marker) || marker == 0x01

        length = blob.byteslice(offset, 2)&.unpack1("n")
        break if length.nil? || length < 2

        segment = blob.byteslice(offset + 2, length - 2)

        if sof?(marker) && segment && segment.bytesize >= 5
          _precision, height, width = segment.unpack("Cnn")
          result[:height] = height
          result[:width] = width
        elsif marker == 0xE0 && segment&.start_with?("JFIF\x00".b)
          apply_jfif_density(result, segment)
        end

        break if marker == 0xDA # start of scan; image data follows

        offset += length
      end
      result
    end

    # SOF0..SOF15, excluding DHT (C4), JPG (C8) and DAC (CC), which share the
    # numeric range but are not frame headers.
    def sof?(marker)
      (0xC0..0xCF).cover?(marker) && ![0xC4, 0xC8, 0xCC].include?(marker)
    end

    def apply_jfif_density(result, segment)
      units, x_density, y_density = segment.byteslice(7, 5).unpack("Cnn")
      case units
      when 1 # dots per inch
        result[:horz_dpi] = x_density
        result[:vert_dpi] = y_density
      when 2 # dots per centimetre
        result[:horz_dpi] = x_density * 2.54
        result[:vert_dpi] = y_density * 2.54
      end
    end

    # GIF carries its dimensions in the logical screen descriptor and has no
    # notion of resolution.
    def parse_gif(blob)
      width, height = blob.byteslice(6, 4).unpack("vv")
      { format: :GIF, width: width, height: height }
    end

    # BMP keeps dimensions and pixels-per-metre in its DIB header. Note that
    # writers generally fill in 3780 ppm rather than leaving it zero, so a BMP
    # usually reports about 96 dpi rather than the 72 default.
    def parse_bmp(blob)
      width, height = blob.byteslice(18, 8).unpack("l<l<")
      ppm_x, ppm_y = blob.byteslice(38, 8).unpack("l<l<")
      result = { format: :BMP, width: width, height: height.abs }
      if ppm_x&.positive? && ppm_y&.positive?
        result[:horz_dpi] = ppm_x * METRES_PER_INCH
        result[:vert_dpi] = ppm_y * METRES_PER_INCH
      end
      result
    end

    TIFF_TAGS = { width: 256, height: 257, x_res: 282, y_res: 283, res_unit: 296 }.freeze

    # TIFF stores everything in tagged IFD entries, in either byte order.
    def parse_tiff(blob)
      little_endian = blob.start_with?("II".b)
      long = little_endian ? "V" : "N"
      short = little_endian ? "v" : "n"

      ifd_offset = blob.byteslice(4, 4).unpack1(long)
      entry_count = blob.byteslice(ifd_offset, 2).unpack1(short)
      values = {}

      entry_count.times do |i|
        entry = blob.byteslice(ifd_offset + 2 + (i * 12), 12)
        break if entry.nil? || entry.bytesize < 12

        tag = entry.byteslice(0, 2).unpack1(short)
        type = entry.byteslice(2, 2).unpack1(short)
        values[tag] = tiff_value(blob, entry, type, long, short)
      end

      result = {
        format: :TIFF,
        width: values[TIFF_TAGS[:width]],
        height: values[TIFF_TAGS[:height]]
      }
      apply_tiff_resolution(result, values)
      result
    end

    # A value of four bytes or fewer is stored inline; anything larger, such as
    # the two longs of a RATIONAL, is stored at an offset.
    def tiff_value(blob, entry, type, long, short)
      payload = entry.byteslice(8, 4)
      case type
      when 3 then payload.unpack1(short)      # SHORT
      when 4 then payload.unpack1(long)       # LONG
      when 5 # RATIONAL
        offset = payload.unpack1(long)
        numerator, denominator = blob.byteslice(offset, 8).unpack("#{long}#{long}")
        denominator.to_i.zero? ? nil : numerator.to_f / denominator
      end
    end

    def apply_tiff_resolution(result, values)
      unit = values[TIFF_TAGS[:res_unit]] || 2
      x_res = values[TIFF_TAGS[:x_res]]
      y_res = values[TIFF_TAGS[:y_res]]
      return if x_res.nil? || y_res.nil?

      case unit
      when 2 # inch
        result[:horz_dpi] = x_res
        result[:vert_dpi] = y_res
      when 3 # centimetre
        result[:horz_dpi] = x_res * 2.54
        result[:vert_dpi] = y_res * 2.54
      end
    end
  end
end
