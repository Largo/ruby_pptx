# frozen_string_literal: true

require "pptx/element_proxy"
require "pptx/shapes/shape_tree"

module Pptx
  # Behaviour common to slides, layouts, masters and notes slides.
  class BaseSlide < PartElementProxy
    # The internal name of this slide; an empty string when unnamed.
    def name = @element.cSld.name

    def name=(value)
      @element.cSld.name = value.to_s
      value
    end

    def shape_tree = @element.spTree
  end

  # Common to slide masters and the notes master.
  class BaseMaster < BaseSlide
    # The shapes on this master.
    def shapes = @shapes ||= MasterShapes.new(@element.spTree, self)

    # The placeholders on this master, in `idx` order.
    def placeholders = @placeholders ||= MasterPlaceholders.new(@element.spTree, self)
  end

  # One slide in a presentation.
  class Slide < BaseSlide
    # The id that identifies this slide within the presentation, stable across
    # reordering.
    def slide_id = part.slide_id

    # The layout this slide takes its appearance from.
    def layout = part.slide_layout

    # True when this slide inherits the master's background.
    def follows_master_background? = @element.bg.nil?

    def has_notes_slide? = part.notes_slide?

    # The shapes on this slide, in z-order.
    def shapes = @shapes ||= SlideShapes.new(@element.spTree, self)

    # The placeholders on this slide, keyed by `idx`.
    def placeholders = @placeholders ||= SlidePlaceholders.new(@element.spTree, self)

    def inspect = "#<Pptx::Slide id=#{slide_id} #{part.partname}>"
  end

  # The slides of a presentation, in order.
  #
  # Enumerable, and indexable by position.
  class Slides < ParentedElementProxy
    include Enumerable

    def initialize(sld_id_list, presentation)
      super(sld_id_list, presentation)
      @sld_id_list = sld_id_list
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @sld_id_list.sldId_list.each { |sld_id| yield part.related_slide(sld_id.rId) }
      self
    end

    def size = @sld_id_list.size
    alias length size

    def empty? = size.zero?

    # Indexed access, supporting a negative index as Ruby arrays do.
    #
    # @return [Slide, nil] nil when +index+ is out of range
    def [](index)
      entry = @sld_id_list.sldId_list[index]
      entry && part.related_slide(entry.rId)
    end

    # As {#[]}, but raises rather than returning nil.
    def fetch(index)
      self[index] || raise(IndexError, "slide index #{index} out of range")
    end

    # The slide with the given slide id.
    #
    # @return [Slide, nil]
    def by_id(slide_id) = part.slide_by_id(slide_id)

    # Add a slide inheriting from +slide_layout+, and return it.
    #
    # The layout's placeholders are copied onto the new slide, preserving
    # z-order; latent ones are left to the layout.
    def add(slide_layout)
      r_id, slide = part.add_slide(slide_layout)
      slide.shapes.clone_layout_placeholders(slide_layout)
      @sld_id_list.add_slide_id(r_id)
      slide
    end

    # The zero-based position of +slide+.
    #
    # @return [Integer, nil] nil when the slide is not in this collection
    def index(slide) = each_with_index.find { |s, _| s == slide }&.last

    def inspect = "#<Pptx::Slides size=#{size}>"
  end

  # A slide layout: the arrangement a slide inherits from.
  class SlideLayout < BaseSlide
    # Placeholders PowerPoint renders from the layout rather than copying onto
    # each slide, so they are not cloned when a slide is created.
    LATENT_PLACEHOLDER_TYPES = [
      Enum::PP_PLACEHOLDER::DATE,
      Enum::PP_PLACEHOLDER::FOOTER,
      Enum::PP_PLACEHOLDER::SLIDE_NUMBER
    ].freeze

    # The shapes on this layout.
    def shapes = @shapes ||= LayoutShapes.new(@element.spTree, self)

    # The placeholders on this layout, in `idx` order.
    def placeholders = @placeholders ||= LayoutPlaceholders.new(@element.spTree, self)

    # The placeholders a new slide based on this layout should receive.
    def cloneable_placeholders
      placeholders.reject { |ph| LATENT_PLACEHOLDER_TYPES.include?(ph.element.ph_type) }
    end

    # The master this layout inherits from.
    def slide_master = part.slide_master

    # The slides based on this layout.
    def used_by_slides
      part.package.presentation_part.presentation.slides.select { |s| s.layout == self }
    end

    def inspect = "#<Pptx::SlideLayout #{name.inspect}>"
  end

  # The layouts belonging to one slide master.
  class SlideLayouts < ParentedElementProxy
    include Enumerable

    def initialize(sld_layout_id_list, slide_master)
      super(sld_layout_id_list, slide_master)
      @sld_layout_id_list = sld_layout_id_list
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @sld_layout_id_list.sldLayoutId_list.each do |entry|
        yield part.related_slide_layout(entry.rId)
      end
      self
    end

    def size = @sld_layout_id_list.size
    alias length size

    # Indexed by position, or looked up by layout name.
    #
    #   layouts[1]
    #   layouts["Title and Content"]
    def [](key)
      return by_name(key) if key.is_a?(String)

      entry = @sld_layout_id_list.sldLayoutId_list[key]
      entry && part.related_slide_layout(entry.rId)
    end

    def fetch(key)
      self[key] || raise(IndexError, "no slide layout #{key.inspect}")
    end

    # @return [SlideLayout, nil]
    def by_name(name) = find { |layout| layout.name == name }

    # @return [Integer, nil]
    def index(slide_layout) = each_with_index.find { |l, _| l == slide_layout }&.last

    # Remove +slide_layout+ from this collection and from the package.
    #
    # @raise [Error] when one or more slides still use the layout
    def delete(slide_layout)
      unless slide_layout.used_by_slides.empty?
        raise Error, "cannot remove a slide layout in use by one or more slides"
      end

      target_index = index(slide_layout) or
        raise NotFoundError, "layout is not in this collection"

      entry = @sld_layout_id_list.sldLayoutId_list[target_index]
      r_id = entry.rId
      # Removing the id stops the layout appearing; dropping the relationship
      # is what actually removes it from the package, along with anything only
      # it referred to.
      @sld_layout_id_list.remove(entry)
      slide_layout.slide_master.part.drop_rel(r_id)
      slide_layout
    end

    def inspect = "#<Pptx::SlideLayouts size=#{size}>"
  end

  # A slide master.
  class SlideMaster < BaseMaster
    # The layouts that inherit from this master.
    def slide_layouts
      @slide_layouts ||= SlideLayouts.new(@element.get_or_add_sldLayoutIdLst, self)
    end

    def inspect = "#<Pptx::SlideMaster #{part.partname}>"
  end

  # The slide masters of a presentation.
  class SlideMasters < ParentedElementProxy
    include Enumerable

    def initialize(sld_master_id_list, presentation)
      super(sld_master_id_list, presentation)
      @sld_master_id_list = sld_master_id_list
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @sld_master_id_list.sldMasterId_list.each do |entry|
        yield part.related_slide_master(entry.rId)
      end
      self
    end

    def size = @sld_master_id_list.size
    alias length size

    def [](index)
      entry = @sld_master_id_list.sldMasterId_list[index]
      entry && part.related_slide_master(entry.rId)
    end

    def inspect = "#<Pptx::SlideMasters size=#{size}>"
  end
end
