# frozen_string_literal: true

require "digest"
require "pptx/opc/constants"

module Pptx
  # A video, as a value object over its bytes.
  #
  # Unlike an image, a video is not interrogated for its type: the caller says
  # what it is. Reading container formats to find out would be a project of its
  # own, and PowerPoint only needs the MIME type recorded correctly.
  class Video
    # What PowerPoint accepts when the type is not known. It plays anyway in
    # practice, which is why python-pptx defaults to it.
    UNKNOWN_CONTENT_TYPE = Opc::CONTENT_TYPE::VIDEO

    EXT_FOR_CONTENT_TYPE = {
      Opc::CONTENT_TYPE::ASF => "asf",
      Opc::CONTENT_TYPE::AVI => "avi",
      Opc::CONTENT_TYPE::MOV => "mov",
      Opc::CONTENT_TYPE::MP4 => "mp4",
      Opc::CONTENT_TYPE::MPG => "mpg",
      Opc::CONTENT_TYPE::MS_VIDEO => "avi",
      Opc::CONTENT_TYPE::SWF => "swf",
      Opc::CONTENT_TYPE::WMV => "wmv",
      Opc::CONTENT_TYPE::X_MS_VIDEO => "avi"
    }.freeze

    attr_reader :blob, :content_type

    def self.from_blob(blob, content_type, filename = nil) = new(blob, content_type, filename)

    # Load from a path or an IO stream.
    def self.from_file(movie_file, content_type)
      if movie_file.is_a?(String)
        new(File.binread(movie_file), content_type, File.basename(movie_file))
      else
        movie_file.rewind if movie_file.respond_to?(:rewind)
        new(movie_file.read, content_type, nil)
      end
    end

    def initialize(blob, content_type, filename = nil)
      @blob = blob
      @content_type = content_type || UNKNOWN_CONTENT_TYPE
      @source_filename = filename
    end

    # The extension for the media part: taken from the original filename when
    # there is one, otherwise derived from the MIME type.
    def ext
      return File.extname(@source_filename).delete_prefix(".") if @source_filename

      EXT_FOR_CONTENT_TYPE.fetch(@content_type, "vid")
    end

    # The name PowerPoint shows for the shape.
    def filename = @source_filename || "movie.#{ext}"

    def sha1 = @sha1 ||= Digest::SHA1.hexdigest(@blob)

    def inspect = "#<Pptx::Video #{filename} #{@content_type}>"
  end
end
