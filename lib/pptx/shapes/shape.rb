# frozen_string_literal: true

require "pptx/shapes/base"
require "pptx/text/text"
require "pptx/dml/fill"
require "pptx/table"

module Pptx
  # A `p:sp`: an auto shape, a text box, or a placeholder.
  class Shape < BaseShape
    def text_frame? = true

    # The text inside this shape, creating an empty text body if it has none.
    def text_frame
      @text_frame ||= TextFrame.new(@element.get_or_add_txBody, self)
    end

    # Shortcut for `text_frame.text`.
    def text = text_frame.text

    def text=(value)
      text_frame.text = value
      value
    end

    # The shape's fill.
    def fill = @fill ||= FillFormat.from_fill_parent(@element.spPr)

    # The shape's outline.
    def line = @line ||= LineFormat.new(@element.spPr)

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

    # The table inside this frame.
    #
    # @raise [Error] when the frame holds something other than a table
    def table
      raise Error, "this graphic frame does not contain a table" unless table?

      @table ||= Table.new(@element.tbl, self)
    end

    # The chart inside this frame.
    #
    # @raise [Error] when the frame holds something other than a chart
    def chart
      raise Error, "this graphic frame does not contain a chart" unless chart?

      @chart ||= part.related_part(@element.chart_rId).chart
    end

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
