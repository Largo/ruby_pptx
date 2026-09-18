# frozen_string_literal: true

require "ruby_pptx/opc/package"
require "ruby_pptx/media"

module Pptx
  module Parts
    # A media part, `/ppt/media/mediaN.<ext>`: the audio or video itself.
    class MediaPart < Opc::Part
      def self.new_media(package, video)
        new(package.next_media_partname(video.ext), video.content_type, package, video.blob)
      end

      def sha1 = @sha1 ||= Digest::SHA1.hexdigest(blob)
    end
  end
end
