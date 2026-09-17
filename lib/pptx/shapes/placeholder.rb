# frozen_string_literal: true

require "pptx/shapes/shape"

module Pptx
  # A placeholder on a slide.
  #
  # Position, size and formatting are inherited from the matching placeholder
  # on the layout unless this shape overrides them, which is why the position
  # accessors can return nil.
  class SlidePlaceholder < Shape
    def shape_type = Enum::MSO_SHAPE_TYPE::PLACEHOLDER

    def inspect
      "#<#{self.class.name} idx=#{placeholder_format.idx} type=#{placeholder_format.type}>"
    end
  end

  # A placeholder on a slide layout.
  class LayoutPlaceholder < Shape
    def shape_type = Enum::MSO_SHAPE_TYPE::PLACEHOLDER
  end

  # A placeholder on a slide master.
  class MasterPlaceholder < Shape
    def shape_type = Enum::MSO_SHAPE_TYPE::PLACEHOLDER
  end
end
