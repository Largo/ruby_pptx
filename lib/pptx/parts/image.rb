# frozen_string_literal: true

require "pptx/opc/package"
require "pptx/image"

module Pptx
  module Parts
    # An image part, `/ppt/media/imageN.<ext>`.
    class ImagePart < Opc::Part
      attr_reader :source_filename

      # A new image part holding +image+.
      def self.new_image(package, image)
        new(package.next_image_partname(image.ext), image.content_type, package,
            image.blob, image.filename)
      end

      def initialize(partname, content_type, package, blob = nil, source_filename = nil)
        super(partname, content_type, package, blob)
        @source_filename = source_filename
      end

      def image = @image ||= Image.from_blob(blob, @source_filename)

      def ext = partname.ext

      # The name PowerPoint shows in the alt-text/description field: the file
      # the image came from, or a generic name when it came from a stream.
      def desc = @source_filename || "image.#{ext}"

      def sha1 = image.sha1

      # Resolve a requested size against the image's native size.
      #
      # Supplying neither dimension uses the native size; supplying one scales
      # the other to preserve the aspect ratio; supplying both stretches the
      # image to fit.
      #
      # @return [Array(Pptx::Length, Pptx::Length)]
      def scale(width, height)
        native_width, native_height = image.native_size
        return [Length.coerce(width), Length.coerce(height)] if width && height
        return [native_width, native_height] unless width || height

        if width
          factor = Length.coerce(width).emu.to_f / native_width.emu
          [Length.coerce(width), Length.emu((native_height.emu * factor).round)]
        else
          factor = Length.coerce(height).emu.to_f / native_height.emu
          [Length.emu((native_width.emu * factor).round), Length.coerce(height)]
        end
      end
    end
  end
end
