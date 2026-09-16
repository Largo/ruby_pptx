# frozen_string_literal: true

module Pptx
  # Base for the objects that make up the public API.
  #
  # Almost every class a caller touches is a thin proxy whose state lives
  # entirely in the XML element it wraps. Two proxies wrapping the same element
  # are equal, whether or not they are the same object.
  class ElementProxy
    attr_reader :element

    def initialize(element)
      @element = element
    end

    def ==(other) = other.is_a?(ElementProxy) && other.element == @element
    alias eql? ==

    def hash = @element.hash

    def inspect = "#<#{self.class.name} <#{@element.nsptag}>>"
  end

  # A proxy that knows its parent, and through it the part it belongs to.
  #
  # An ancestor is occasionally needed to do something the proxy cannot, such
  # as add or drop a relationship.
  class ParentedElementProxy < ElementProxy
    attr_reader :parent

    def initialize(element, parent)
      super(element)
      @parent = parent
    end

    # The package part this object lives in.
    def part = @parent.part
  end

  # A proxy wrapping a part's root element, such as `p:sld`.
  class PartElementProxy < ElementProxy
    attr_reader :part

    def initialize(element, part)
      super(element)
      @part = part
    end
  end
end
