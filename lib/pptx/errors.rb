# frozen_string_literal: true

module Pptx
  # Base class for every error raised by this library.
  class Error < StandardError; end

  # Raised when a file is not a readable OOXML package.
  class PackageNotFoundError < Error; end

  # Raised when the XML in a package part violates the schema in a way we rely on.
  class InvalidXmlError < Error; end

  # Raised when a requested shape, part or relationship does not exist.
  class NotFoundError < Error; end
end
