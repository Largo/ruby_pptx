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

    # The adjustment handles of this shape's geometry; empty for a freeform.
    #
    # @return [Adjustments]
    def adjustments = @adjustments ||= Adjustments.new(@element.spPr.prstGeom)

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
  class BasePicture < BaseShape
    # Cropping, as the fraction of the image cut from each side: 0.25 is a
    # quarter. Negative values extend the side past the image edge, and values
    # above 1.0 are allowed, as in PowerPoint.
    %i[left top right bottom].each do |side|
      attribute = :"srcRect_#{side.to_s[0]}"
      define_method(:"crop_#{side}") { @element.public_send(attribute) }
      define_method(:"crop_#{side}=") { |value| @element.public_send(:"#{attribute}=", value) }
    end

    # The outline drawn around the picture.
    def line = @line ||= LineFormat.new(@element.spPr)
  end

  # A `p:pic` showing a still image.
  class Picture < BasePicture
    def shape_type = Enum::MSO_SHAPE_TYPE::PICTURE

    # The image this picture shows: its bytes, format, size and resolution.
    #
    # @return [Pptx::Image]
    # @raise [Error] when the picture links its image rather than embedding it
    def image
      r_id = @element.blip_rId
      raise Error, "this picture has no embedded image" if r_id.nil?

      part.related_part(r_id).image
    end

    # The auto shape the picture is masked by. A new picture is a rectangle,
    # which crops nothing; an ellipse shows it through an oval.
    #
    # @return [Pptx::Enum::Member, nil] nil for a freeform mask, or no geometry
    def auto_shape_type = @element.spPr.prstGeom&.prst

    def auto_shape_type=(value)
      member = Enum::MSO_AUTO_SHAPE_TYPE.fetch(value)
      sp_pr = @element.spPr
      geometry = sp_pr.prstGeom
      if geometry.nil?
        sp_pr.remove_custGeom
        geometry = sp_pr.get_or_add_prstGeom
      end
      geometry.prst = member
    end
  end

  # A `p:pic` that carries video rather than a still image.
  #
  # Not a {Picture}: the image a movie holds is its poster frame, which is
  # what {#poster_frame} returns, and a movie cannot be masked by a shape.
  class Movie < BasePicture
    def shape_type = Enum::MSO_SHAPE_TYPE::MEDIA

    def media_type = Enum::PP_MEDIA_TYPE::MOVIE

    # The still shown before the movie plays, or nil if there is none.
    #
    # @return [Pptx::Image, nil]
    def poster_frame
      r_id = @element.blip_rId
      r_id && part.related_part(r_id).image
    end

    # Playback settings. PowerPoint keeps these in the timing tree and
    # python-pptx exposes none of them yet; the object exists so code written
    # against either library finds it.
    def media_format = @media_format ||= MediaFormat.new(@element)
  end

  # Playback formatting for a movie. Deliberately empty; see {Movie#media_format}.
  class MediaFormat < ElementProxy; end

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
    # A graphic frame's shadow belongs to what it holds -- a chart and a
    # table keep theirs in different places -- so there is no single one to
    # hand back. python-pptx declines the same way.
    def shadow
      raise Error, "a graphic frame has no shadow of its own; format the chart or table instead"
    end

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

    # An OLE object is linked when it points at a file outside the
    # presentation rather than carrying one. nil for anything else a frame
    # can hold, such as a diagram, as in python-pptx.
    def shape_type
      return Enum::MSO_SHAPE_TYPE::TABLE if table?
      return Enum::MSO_SHAPE_TYPE::CHART if chart?
      return nil unless ole_object?

      if @element.embedded_ole_object?
        Enum::MSO_SHAPE_TYPE::EMBEDDED_OLE_OBJECT
      else
        Enum::MSO_SHAPE_TYPE::LINKED_OLE_OBJECT
      end
    end

    def ole_object? = @element.ole_object?

    # The OLE object this frame holds.
    #
    # @raise [Error] when the frame holds something else
    def ole_format
      raise Error, "this graphic frame does not contain an OLE object" unless ole_object?

      @ole_format ||= OleFormat.new(@element.oleObj, self)
    end
  end

  # An embedded or linked OLE object: the file, and how it is shown.
  class OleFormat < ElementProxy
    def initialize(ole_obj, frame)
      super(ole_obj)
      @frame = frame
    end

    # The embedded file's bytes, or nil for a linked object.
    def blob
      r_id = @element.rId
      r_id && @frame.part.related_part(r_id).blob
    end

    # The ProgID naming the program that opens the object, e.g. "Excel.Sheet.12".
    def prog_id = @element.progId

    # Whether it appears as an icon rather than as a picture of its content.
    def show_as_icon? = @element.showAsIcon

    def inspect = "#<Pptx::OleFormat #{prog_id.inspect}>"
  end

  # A `p:grpSp` containing other shapes.
  #
  # A group has no position or size of its own: both follow from what it
  # contains, and are recomputed whenever its contents change.
  class GroupShape < BaseShape
    # A group keeps its effects in `p:grpSpPr` rather than `p:spPr`.
    def shadow = @shadow ||= ShadowFormat.new(@element.grpSpPr)

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
