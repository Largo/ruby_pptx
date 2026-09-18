# frozen_string_literal: true

require "ruby_pptx/shapes/shape"

module Pptx
  # Position and size that fall back to the placeholder being inherited from.
  #
  # A placeholder usually carries no `a:xfrm` of its own: it takes its geometry
  # from the layout, which takes it from the master. Reading it back therefore
  # has to walk that chain, or every placeholder on a freshly built slide looks
  # like it has no position at all.
  #
  # Writing is not inherited -- setting a value applies it to this shape, which
  # is what stops it tracking the layout from then on.
  module InheritsDimensions
    def left = effective(:left)
    def top = effective(:top)
    def width = effective(:width)
    def height = effective(:height)

    # The placeholder one level up that this one inherits from, or nil.
    def base_placeholder = raise(NotImplementedError, "#{self.class} must define base_placeholder")

    private

    def effective(attribute)
      own = super_value(attribute)
      return own unless own.nil?

      base_placeholder&.public_send(attribute)
    end

    def super_value(attribute)
      method(attribute).super_method.call
    end
  end

  # A placeholder on a slide.
  #
  # Position, size and formatting are inherited from the matching placeholder
  # on the layout unless this shape overrides them.
  class SlidePlaceholder < Shape
    include InheritsDimensions

    def shape_type = Enum::MSO_SHAPE_TYPE::PLACEHOLDER

    # Matched to the layout by `idx`, which is what ties the two together.
    def base_placeholder
      part.slide_layout.placeholders.by_idx(element.ph_idx)
    end

    def inspect
      "#<#{self.class.name} idx=#{placeholder_format.idx} type=#{placeholder_format.type}>"
    end
  end

  # A placeholder on a slide layout.
  class LayoutPlaceholder < Shape
    include InheritsDimensions

    # A layout placeholder inherits from the master placeholder of the
    # corresponding type rather than the same one: a chart, table or picture
    # placeholder all take their geometry from the master's body.
    BASE_TYPES = {
      BODY: :BODY, CHART: :BODY, BITMAP: :BODY, ORG_CHART: :BODY, MEDIA_CLIP: :BODY,
      OBJECT: :BODY, PICTURE: :BODY, SUBTITLE: :BODY, TABLE: :BODY,
      CENTER_TITLE: :TITLE, TITLE: :TITLE,
      DATE: :DATE, FOOTER: :FOOTER, SLIDE_NUMBER: :SLIDE_NUMBER
    }.freeze

    def shape_type = Enum::MSO_SHAPE_TYPE::PLACEHOLDER

    def base_placeholder
      base_type = BASE_TYPES[element.ph_type&.name]
      return nil if base_type.nil?

      part.slide_master.placeholders.by_type(Enum::PP_PLACEHOLDER_TYPE.fetch(base_type))
    end
  end

  # A placeholder on a slide master.
  #
  # There is nothing above a master to inherit from, so its geometry is
  # whatever it carries itself.
  class MasterPlaceholder < Shape
    def shape_type = Enum::MSO_SHAPE_TYPE::PLACEHOLDER
  end
end
