# frozen_string_literal: true

module Pptx
  # The shadow of a shape.
  #
  # Only inheritance can be controlled so far, as in python-pptx: a shape
  # either takes its shadow from the theme and style hierarchy, or has its
  # effects switched off explicitly.
  #
  #   shape.shadow.inherit = false   # no shadow, whatever the theme says
  class ShadowFormat
    # @param shape_properties [Pptx::Oxml::Element] `p:spPr` or `p:grpSpPr`,
    #   which both carry the `a:effectLst`
    def initialize(shape_properties)
      @element = shape_properties
    end

    # True when the shape takes its shadow from the style hierarchy.
    def inherit?
      @element.effectLst.nil?
    end

    alias inherit inherit?

    # Restore inheritance with true; with false, break it so no effects show.
    #
    # Inheritance is all or nothing: restoring it removes every explicit
    # effect on the shape -- glow and reflection too, not only the shadow.
    def inherit=(value)
      value ? @element.remove_effectLst : @element.get_or_add_effectLst
    end

    def inspect
      "#<Pptx::ShadowFormat inherit=#{inherit?}>"
    end
  end
end
