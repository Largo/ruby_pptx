# frozen_string_literal: true

require "pptx/element_proxy"
require "pptx/shapes/base"
require "pptx/shapes/shape"
require "pptx/shapes/placeholder"
require "pptx/autoshape_spec"

module Pptx
  # Builds the right shape proxy for a shape element.
  module ShapeFactory
    module_function

    # The default mapping, used for shapes on a slide master or layout and as
    # the base for the slide-specific variants.
    def build(shape_element, parent)
      case shape_element.nsptag
      when "p:pic"
        shape_element.movie? ? Movie.new(shape_element, parent) : Picture.new(shape_element, parent)
      when "p:cxnSp" then Connector.new(shape_element, parent)
      when "p:grpSp" then GroupShape.new(shape_element, parent)
      when "p:sp" then Shape.new(shape_element, parent)
      when "p:graphicFrame" then GraphicFrame.new(shape_element, parent)
      else BaseShape.new(shape_element, parent)
      end
    end

    # On a slide, a `p:sp` that is a placeholder gets the placeholder proxy.
    def build_for_slide(shape_element, parent)
      return SlidePlaceholder.new(shape_element, parent) if placeholder_sp?(shape_element)

      build(shape_element, parent)
    end

    def build_for_layout(shape_element, parent)
      return LayoutPlaceholder.new(shape_element, parent) if placeholder_sp?(shape_element)

      build(shape_element, parent)
    end

    def build_for_master(shape_element, parent)
      return MasterPlaceholder.new(shape_element, parent) if placeholder_sp?(shape_element)

      build(shape_element, parent)
    end

    def placeholder_sp?(shape_element)
      shape_element.nsptag == "p:sp" && shape_element.placeholder?
    end
  end

  # The shapes of a slide, layout or master, in z-order: first is backmost.
  class BaseShapes < ParentedElementProxy
    include Enumerable

    def initialize(sp_tree, parent)
      super(sp_tree, parent)
      @sp_tree = sp_tree
    end

    def each
      return enum_for(:each) { size } unless block_given?

      member_elements.each { |element| yield shape_factory(element) }
      self
    end

    # A group counts as one shape, whatever it contains.
    def size = member_elements.size
    alias length size

    def [](index)
      element = member_elements[index]
      element && shape_factory(element)
    end

    def fetch(index)
      self[index] || raise(IndexError, "shape index #{index} out of range")
    end

    # Add a placeholder to this collection modelled on +placeholder+.
    def clone_placeholder(placeholder)
      sp = placeholder.element
      id = next_shape_id
      name = next_placeholder_name(sp.ph_type, id, sp.ph_orient)
      @sp_tree.add_placeholder(id, name, sp.ph_type, sp.ph_orient, sp.ph_sz, sp.ph_idx)
    end

    # The base name PowerPoint gives a placeholder of +ph_type+.
    #
    # A notes slide names its body placeholder differently, so subclasses can
    # override this.
    def placeholder_basename(ph_type)
      PLACEHOLDER_BASENAMES.fetch(ph_type.name) do
        raise ArgumentError, "no placeholder base name for #{ph_type}"
      end
    end

    PLACEHOLDER_BASENAMES = {
      BITMAP: "ClipArt Placeholder",
      BODY: "Text Placeholder",
      CENTER_TITLE: "Title",
      CHART: "Chart Placeholder",
      DATE: "Date Placeholder",
      FOOTER: "Footer Placeholder",
      HEADER: "Header Placeholder",
      MEDIA_CLIP: "Media Placeholder",
      OBJECT: "Content Placeholder",
      ORG_CHART: "SmartArt Placeholder",
      PICTURE: "Picture Placeholder",
      SLIDE_NUMBER: "Slide Number Placeholder",
      SUBTITLE: "Subtitle",
      TABLE: "Table Placeholder",
      TITLE: "Title"
    }.freeze

    private

    # Which shape elements belong to this collection; placeholder collections
    # narrow this.
    def member?(_shape_element) = true

    def member_elements = @sp_tree.shape_elements.select { |e| member?(e) }

    def shape_factory(shape_element) = ShapeFactory.build(shape_element, self)

    # One more than the highest id in use. Note the shape tree's own allocator
    # fills gaps instead; both behaviours are inherited from python-pptx.
    def next_shape_id = @sp_tree.max_shape_id + 1

    # The name PowerPoint would give a new placeholder: the base name for its
    # type followed by id - 1, bumped until it is unique in the tree, and
    # prefixed "Vertical " for a vertical placeholder.
    def next_placeholder_name(ph_type, id, orient)
      basename = placeholder_basename(ph_type)
      basename = "Vertical #{basename}" if orient == Oxml::SimpleTypes::ST_Direction::VERT

      taken = @sp_tree.xpath("//p:cNvPr/@name").map(&:value)
      numpart = id - 1
      numpart += 1 while taken.include?("#{basename} #{numpart}")
      "#{basename} #{numpart}"
    end
  end

  # The shapes on a slide.
  class SlideShapes < BaseShapes
    # Copy the layout's placeholders onto this slide, preserving z-order.
    #
    # Latent placeholders -- date, footer and slide number -- are not cloned:
    # PowerPoint shows them from the layout without materialising them on the
    # slide.
    def clone_layout_placeholders(slide_layout)
      slide_layout.cloneable_placeholders.each { |ph| clone_placeholder(ph) }
      self
    end

    # The title placeholder, which is the one with idx 0.
    #
    # @return [SlidePlaceholder, nil]
    def title
      element = @sp_tree.placeholder_elements.find { |e| e.ph_idx.zero? }
      element && shape_factory(element)
    end

    def placeholders = parent.placeholders

    # Add an auto shape of +shape_type+ at the given position and size.
    #
    #   shapes.add_shape(Pptx::Enum::MSO_SHAPE::ROUNDED_RECTANGLE,
    #                    Pptx.inches(1), Pptx.inches(1),
    #                    Pptx.inches(2), Pptx.inches(1))
    #
    # @param shape_type [Pptx::Enum::MSO_SHAPE, Symbol, Integer]
    # @return [Shape]
    def add_shape(shape_type, left, top, width, height)
      member = Enum::MSO_SHAPE.fetch(shape_type)
      id = next_shape_id
      name = "#{AutoShapeSpec.basename(member)} #{id - 1}"
      sp = @sp_tree.add_autoshape(id, name, Enum::MSO_SHAPE.to_xml(member),
                                  left, top, width, height)
      shape_factory(sp)
    end

    # Add a picture showing the image in +image_file+, which may be a path or
    # an IO stream.
    #
    # Supplying neither +width+ nor +height+ uses the image's native size;
    # supplying one scales the other to preserve the aspect ratio; supplying
    # both stretches the image to fit.
    #
    # @return [Picture]
    def add_picture(image_file, left, top, width: nil, height: nil)
      image_part, r_id = part.get_or_add_image_part(image_file)
      scaled_width, scaled_height = image_part.scale(width, height)
      id = next_shape_id
      pic = @sp_tree.add_pic(id, "Picture #{id - 1}", image_part.desc, r_id,
                             left, top, scaled_width, scaled_height)
      shape_factory(pic)
    end

    # Add a table of +rows+ by +cols+ filling the given position and size.
    #
    # @return [GraphicFrame] use its `#table` to reach the table itself
    def add_table(rows, cols, left, top, width, height)
      id = next_shape_id
      frame = @sp_tree.add_graphic_frame_table(id, "Table #{id - 1}", rows, cols,
                                               left, top, width, height)
      shape_factory(frame)
    end

    # Add an empty text box at the given position and size.
    #
    # @return [Shape]
    def add_textbox(left, top, width, height)
      id = next_shape_id
      sp = @sp_tree.add_textbox(id, "TextBox #{id - 1}", left, top, width, height)
      shape_factory(sp)
    end

    private

    def shape_factory(shape_element) = ShapeFactory.build_for_slide(shape_element, self)
  end

  # The shapes inside a `p:grpSp`.
  class GroupShapes < BaseShapes
  end

  # The shapes on a slide layout.
  class LayoutShapes < BaseShapes
    private

    def shape_factory(shape_element) = ShapeFactory.build_for_layout(shape_element, self)
  end

  # The shapes on a slide master.
  class MasterShapes < BaseShapes
    private

    def shape_factory(shape_element) = ShapeFactory.build_for_master(shape_element, self)
  end

  # The placeholders of a slide layout, in `idx` order.
  class LayoutPlaceholders < LayoutShapes
    # @return [LayoutPlaceholder, nil]
    def by_idx(idx) = find { |ph| ph.element.ph_idx == idx }

    private

    def member?(shape_element) = shape_element.placeholder?

    def member_elements = super.sort_by(&:ph_idx)
  end

  # The placeholders of a slide master, in `idx` order.
  class MasterPlaceholders < MasterShapes
    # @return [MasterPlaceholder, nil]
    def by_type(ph_type) = find { |ph| ph.element.ph_type == ph_type }

    private

    def member?(shape_element) = shape_element.placeholder?

    def member_elements = super.sort_by(&:ph_idx)
  end

  # The placeholders of a slide.
  #
  # Ordered by `idx` and looked up by it: `placeholders[1]` is the placeholder
  # whose idx is 1, not the second one in the collection.
  class SlidePlaceholders < ParentedElementProxy
    include Enumerable

    def initialize(sp_tree, parent)
      super(sp_tree, parent)
      @sp_tree = sp_tree
    end

    def each
      return enum_for(:each) { size } unless block_given?

      placeholder_elements.each { |e| yield ShapeFactory.build_for_slide(e, self) }
      self
    end

    def size = @sp_tree.placeholder_elements.count
    alias length size

    # The placeholder whose `idx` is +idx+, or nil.
    def [](idx)
      element = @sp_tree.placeholder_elements.find { |e| e.ph_idx == idx }
      element && ShapeFactory.build_for_slide(element, self)
    end

    def fetch(idx)
      self[idx] || raise(NotFoundError, "no placeholder on this slide with idx #{idx}")
    end

    private

    def placeholder_elements = @sp_tree.placeholder_elements.sort_by(&:ph_idx)
  end
end
