# frozen_string_literal: true

require "ruby_pptx/shapes/base"
require "ruby_pptx/text/text"
require "ruby_pptx/dml/fill"
require "ruby_pptx/table"

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

  # A `p:cxnSp`: a line, optionally attached to shapes at either end.
  #
  # A connector is stored as a bounding box plus flip flags rather than as two
  # points, so moving an end point may flip the box rather than move it. The
  # accessors here hide that.
  class Connector < BaseShape
    def shape_type = Enum::MSO_SHAPE_TYPE::LINE

    def line = @line ||= LineFormat.new(@element.spPr)

    def begin_x = Length.emu(@element.flipH ? left.to_i + width.to_i : left.to_i)

    def begin_x=(value)
      move_x(Length.coerce(value).emu, begin_point: true)
    end

    def begin_y = Length.emu(@element.flipV ? top.to_i + height.to_i : top.to_i)

    def begin_y=(value)
      move_y(Length.coerce(value).emu, begin_point: true)
    end

    def end_x = Length.emu(@element.flipH ? left.to_i : left.to_i + width.to_i)

    def end_x=(value)
      move_x(Length.coerce(value).emu, begin_point: false)
    end

    def end_y = Length.emu(@element.flipV ? top.to_i : top.to_i + height.to_i)

    def end_y=(value)
      move_y(Length.coerce(value).emu, begin_point: false)
    end

    # Attach the start of this connector to +shape+ at one of its connection
    # points, and move the start there.
    #
    # Connection points are numbered from 0, conventionally starting at the top
    # centre and going counter-clockwise. That is a convention rather than a
    # rule, so the result is only predictable for rectangular shapes.
    def begin_connect(shape, connection_point)
      connection = @element.nvCxnSpPr.cNvCxnSpPr.get_or_add_stCxn
      connection.id = shape.shape_id
      connection.idx = connection_point
      x, y = connection_point_of(shape, connection_point)
      self.begin_x = x
      self.begin_y = y
      self
    end

    # As {#begin_connect}, for the far end.
    def end_connect(shape, connection_point)
      connection = @element.nvCxnSpPr.cNvCxnSpPr.get_or_add_endCxn
      connection.id = shape.shape_id
      connection.idx = connection_point
      x, y = connection_point_of(shape, connection_point)
      self.end_x = x
      self.end_y = y
      self
    end

    def begin_connected? = !@element.nvCxnSpPr.cNvCxnSpPr.stCxn.nil?

    def end_connected? = !@element.nvCxnSpPr.cNvCxnSpPr.endCxn.nil?

    def inspect
      "#<Pptx::Connector (#{begin_x.inches.round(2)}, #{begin_y.inches.round(2)}) -> " \
        "(#{end_x.inches.round(2)}, #{end_y.inches.round(2)})>"
    end

    private

    # The four points PowerPoint puts on a rectangle: top, left, bottom, right.
    def connection_point_of(shape, index)
      x = shape.left.to_i
      y = shape.top.to_i
      cx = shape.width.to_i
      cy = shape.height.to_i
      case index
      when 0 then [x + (cx / 2), y]
      when 1 then [x, y + (cy / 2)]
      when 2 then [x + (cx / 2), y + cy]
      when 3 then [x + cx, y + (cy / 2)]
      else
        raise ArgumentError, "connection point #{index} is out of range (0-3)"
      end
    end

    # Move one end along an axis, keeping the other end where it is.
    #
    # The box has no notion of direction, so pulling an end past the other one
    # flips the box rather than giving it a negative size.
    def move_x(new_value, begin_point:)
      fixed = begin_point ? end_x.emu : begin_x.emu
      @element.x = [new_value, fixed].min
      @element.cx = (new_value - fixed).abs
      @element.flipH = begin_point ? new_value > fixed : new_value < fixed
    end

    def move_y(new_value, begin_point:)
      fixed = begin_point ? end_y.emu : begin_y.emu
      @element.y = [new_value, fixed].min
      @element.cy = (new_value - fixed).abs
      @element.flipV = begin_point ? new_value > fixed : new_value < fixed
    end
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
  #
  # A group has no position or size of its own: both follow from what it
  # contains, and are recomputed whenever its contents change.
  class GroupShape < BaseShape
    def shape_type = Enum::MSO_SHAPE_TYPE::GROUP

    # The shapes inside this group.
    def shapes = @shapes ||= GroupShapes.new(@element, self)

    # A group holds no text of its own, though the shapes in it may.
    def text_frame? = false

    # PowerPoint offers no click behaviour on a group; the shapes inside carry
    # their own.
    def click_action
      raise Error, "a group shape cannot have a click action"
    end

    def hyperlink
      raise Error, "a group shape cannot have a hyperlink"
    end
  end
end
