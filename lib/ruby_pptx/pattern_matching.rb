# frozen_string_literal: true

module Pptx
  # `case`/`in` support for the objects a caller reads a deck with.
  #
  #   case shape
  #   in {shape_type: :PICTURE, name:}          then puts "picture #{name}"
  #   in {shape_type: :PLACEHOLDER, placeholder_format: {type: :TITLE}}
  #                                             then puts shape.text_frame.text
  #   in {width: Pptx::Length => w} if w > Pptx.inches(5)
  #                                             then puts "wide"
  #   end
  #
  # **Enum-valued attributes read as their symbolic name in a pattern.**
  # `shape.shape_type` returns a {Pptx::Enum::Member}, because that member
  # carries the MS API value and the XML value and is the honest return type.
  # In a pattern none of that helps -- `in {shape_type: :PICTURE}` is what
  # anyone would write -- so the pattern-matching view reports the symbol.
  #
  # Only the keys a pattern actually asks for are computed, so matching on a
  # shape's name does not walk its text.
  module PatternMatching
    def self.included(base)
      base.extend(ClassMethods)
    end

    # A member reads as its name; anything else is passed through.
    def self.for_pattern(value)
      value.is_a?(Enum::Member) ? value.to_sym : value
    end

    module ClassMethods
      # Declare the keys `case`/`in` may match on. Each names a public method.
      def pattern_keys(*names)
        names.freeze
        define_method(:deconstruct_keys) do |keys|
          wanted = keys.nil? ? names : names & keys
          wanted.to_h { |name| [name, PatternMatching.for_pattern(public_send(name))] }
        end
      end
    end
  end

  # Array patterns for the collections, so `in [first, *rest]` works.
  #
  # Kept separate from {PatternMatching} because deconstructing a collection
  # means enumerating all of it, which is the opposite of the lazy wrapping
  # everything else here does. Use it where the whole list is wanted anyway.
  module DeconstructToArray
    def deconstruct
      to_a
    end
  end
end
