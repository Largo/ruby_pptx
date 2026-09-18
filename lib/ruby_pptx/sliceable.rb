# frozen_string_literal: true

module Pptx
  # Array-like indexing for the collections that wrap their members lazily.
  #
  # Every collection in this library holds a list of XML elements and builds
  # the Ruby object for one only when it is asked for. That makes `[]` a little
  # more than a delegation: an integer index yields one wrapped member, while a
  # range or a (start, length) pair yields an array of them.
  #
  #   slides[2]      #=> a Slide
  #   slides[-1]     #=> the last Slide
  #   slides[1..3]   #=> an Array of Slides
  #   slides[1, 2]   #=> an Array of Slides
  #
  # Out of range gives nil, and a range that starts past the end gives nil
  # rather than [], matching Array.
  module Sliceable
    private

    # @param members [Array] the backing elements
    # @yieldparam member [Object] one backing element, to be wrapped
    # @return [Object, Array, nil]
    def slice_members(members, index, length = nil, &)
      if length.nil? && !index.is_a?(Range)
        member = members[index]
        return member && yield(member)
      end

      selected = length.nil? ? members[index] : members[index, length]
      selected&.map(&)
    end
  end
end
