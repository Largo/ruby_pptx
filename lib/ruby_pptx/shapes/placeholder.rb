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
    def left
      effective(:left)
    end

    def top
      effective(:top)
    end

    def width
      effective(:width)
    end

    def height
      effective(:height)
    end

    # The placeholder one level up that this one inherits from, or nil.
    def base_placeholder
      raise(NotImplementedError, "#{self.class} must define base_placeholder")
    end

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

    def shape_type
      Enum::MSO_SHAPE_TYPE::PLACEHOLDER
    end

    # Matched to the layout by `idx`, which is what ties the two together.
    def base_placeholder
      part.slide_layout.placeholders.by_idx(element.ph_idx)
    end

    def inspect
      "#<#{self.class.name} idx=#{placeholder_format.idx} type=#{placeholder_format.type}>"
    end
  end

  # A slide placeholder that can be filled with something other than text:
  # the new content takes over the placeholder's slot, id and name, and keeps
  # the `p:ph` that ties it to the layout.
  module FillablePlaceholder
    private

    # Put +new_element+ where this placeholder is and move the `p:ph` onto
    # it. The original element leaves the tree; this object is spent.
    def replace_placeholder_with(new_element)
      adopted = @element.import(new_element)
      adopted.nvXxPr.nvPr.insert_ph(@element.ph)
      @element.add_previous_sibling(adopted)
      @element.parent.remove(@element)
      ShapeFactory.build_for_slide(adopted, @parent)
    end
  end

  # A picture placeholder: fill it with {#insert_picture}.
  class PicturePlaceholder < SlidePlaceholder
    include FillablePlaceholder

    # Fill this placeholder with an image, cropped to cover it exactly.
    #
    # The picture keeps no size of its own -- it inherits the placeholder's
    # from the layout -- so the image is cropped, never distorted, to fit.
    #
    # @param image_file [String, IO] a path or an open stream
    # @return [PlaceholderPicture]
    def insert_picture(image_file)
      image_part, r_id = part.get_or_add_image_part(image_file)
      pic = Oxml::CT_Picture.new_ph_pic(shape_id, name, image_part.desc, r_id)
      pic.crop_to_fit(image_part.image.size, [width.emu, height.emu])
      replace_placeholder_with(pic)
    end
  end

  # A table placeholder: fill it with {#insert_table}.
  class TablePlaceholder < SlidePlaceholder
    include FillablePlaceholder

    # PowerPoint's default row height, which sets the new table's height.
    ROW_HEIGHT = 370_840

    # Fill this placeholder with an empty table at its position and width.
    #
    # @return [PlaceholderGraphicFrame] whose #table is the new table
    def insert_table(rows, cols)
      frame = Oxml::CT_GraphicalObjectFrame.new_table_graphic_frame(
        shape_id, name, rows, cols, left, top, width, Length.emu(rows * ROW_HEIGHT)
      )
      replace_placeholder_with(frame)
    end
  end

  # A chart placeholder: fill it with {#insert_chart}.
  class ChartPlaceholder < SlidePlaceholder
    include FillablePlaceholder

    # Fill this placeholder with a chart occupying exactly its area.
    #
    # @return [PlaceholderGraphicFrame] whose #chart is the new chart
    def insert_chart(chart_type, chart_data)
      r_id = part.add_chart_part(chart_type, chart_data)
      frame = Oxml::CT_GraphicalObjectFrame.new_chart_graphic_frame(
        shape_id, name, r_id, left, top, width, height
      )
      replace_placeholder_with(frame)
    end
  end

  # A picture that fills a placeholder. It is a {Picture} in every respect
  # except that it inherits its geometry from the layout.
  class PlaceholderPicture < Picture
    include InheritsDimensions

    def shape_type
      Enum::MSO_SHAPE_TYPE::PLACEHOLDER
    end

    def base_placeholder
      part.slide_layout.placeholders.by_idx(element.ph_idx)
    end
  end

  # A table or chart that fills a placeholder.
  class PlaceholderGraphicFrame < GraphicFrame
    def placeholder?
      true
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

    def shape_type
      Enum::MSO_SHAPE_TYPE::PLACEHOLDER
    end

    def base_placeholder
      base_type = BASE_TYPES[element.ph_type&.name]
      return nil if base_type.nil?

      part.slide_master.placeholders.by_type(Enum::PP_PLACEHOLDER_TYPE.fetch(base_type))
    end
  end

  # A placeholder on a notes page.
  #
  # Inherits its geometry from the notes-master placeholder of the same type.
  class NotesSlidePlaceholder < Shape
    include InheritsDimensions

    def shape_type
      Enum::MSO_SHAPE_TYPE::PLACEHOLDER
    end

    def base_placeholder
      part.notes_master.placeholders.by_type(element.ph_type)
    end
  end

  # A placeholder on a slide master.
  #
  # There is nothing above a master to inherit from, so its geometry is
  # whatever it carries itself.
  class MasterPlaceholder < Shape
    def shape_type
      Enum::MSO_SHAPE_TYPE::PLACEHOLDER
    end
  end
end
