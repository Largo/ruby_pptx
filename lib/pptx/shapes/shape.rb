# frozen_string_literal: true

require "pptx/shapes/base"

module Pptx
  # A `p:sp`: an auto shape, a text box, or a placeholder.
  class Shape < BaseShape
    def text_frame? = true

    # The preset geometry of an auto shape.
    #
    # @return [Pptx::Enum::MSO_SHAPE, nil] nil unless this is an auto shape
    def auto_shape_type = @element.autoshape? ? @element.prst : nil

    def shape_type
      return Enum::MSO_SHAPE_TYPE::PLACEHOLDER if placeholder?
      return Enum::MSO_SHAPE_TYPE::TEXT_BOX if @element.textbox?
      return Enum::MSO_SHAPE_TYPE::FREEFORM if @element.custom_geometry?
      return Enum::MSO_SHAPE_TYPE::AUTO_SHAPE if @element.autoshape?

      Enum::MSO_SHAPE_TYPE::AUTO_SHAPE
    end
  end

  # A `p:pic` holding an image.
  class Picture < BaseShape
    def shape_type = Enum::MSO_SHAPE_TYPE::PICTURE
  end

  # A `p:pic` that carries video rather than a still image.
  class Movie < Picture
    def shape_type = Enum::MSO_SHAPE_TYPE::MEDIA
  end

  # A `p:cxnSp` joining two shapes.
  class Connector < BaseShape
    def shape_type = Enum::MSO_SHAPE_TYPE::LINE
  end

  # A `p:graphicFrame`, which holds a table, a chart or an embedded object.
  class GraphicFrame < BaseShape
    def table? = @element.table?
    def chart? = @element.chart?

    def shape_type
      return Enum::MSO_SHAPE_TYPE::TABLE if table?
      return Enum::MSO_SHAPE_TYPE::CHART if chart?

      Enum::MSO_SHAPE_TYPE::EMBEDDED_OLE_OBJECT
    end
  end

  # A `p:grpSp` containing other shapes.
  class GroupShape < BaseShape
    def shape_type = Enum::MSO_SHAPE_TYPE::GROUP

    # The shapes inside this group.
    def shapes = @shapes ||= GroupShapes.new(@element, self)
  end
end
