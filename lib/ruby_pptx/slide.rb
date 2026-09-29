# frozen_string_literal: true

require "ruby_pptx/pattern_matching"

require "ruby_pptx/element_proxy"
require "ruby_pptx/shapes/shape_tree"
require "ruby_pptx/dml/fill"

module Pptx
  # Behaviour common to slides, layouts, masters and notes slides.
  class BaseSlide < PartElementProxy
    # The internal name of this slide; an empty string when unnamed.
    def name
      @element.cSld.name
    end

    def name=(value)
      @element.cSld.name = value.to_s
    end

    def shape_tree
      @element.spTree
    end

    # The slide background.
    #
    # Note that merely reading this is destructive: a background given by a
    # style reference, or inherited, is replaced with an explicit no-fill so
    # there is something to interrogate. python-pptx behaves the same way.
    def background
      @background ||= Background.new(@element.cSld)
    end
  end

  # The background of a slide, layout or master.
  class Background < ElementProxy
    def fill
      @fill ||= FillFormat.from_fill_parent(@element.get_or_add_bgPr)
    end
  end

  # Common to slide masters and the notes master.
  class BaseMaster < BaseSlide
    # The shapes on this master.
    def shapes
      @shapes ||= MasterShapes.new(@element.spTree, self)
    end

    # The placeholders on this master, in `idx` order.
    def placeholders
      @placeholders ||= MasterPlaceholders.new(@element.spTree, self)
    end
  end

  # The master every notes page in a presentation inherits from.
  #
  # There is at most one per presentation, and it is created on first use --
  # the first time a slide is given speaker notes, or {Presentation#notes_master}
  # is asked for.
  class NotesMaster < BaseMaster
    def inspect
      "#<Pptx::NotesMaster #{part.partname}>"
    end
  end

  # The speaker notes page belonging to one slide.
  #
  #   slide.notes_slide.notes_text_frame.text = "Mention the Q3 numbers"
  #   slide.notes = "Mention the Q3 numbers"   # the same, shorter
  class NotesSlide < BaseSlide
    # Placeholder types copied from the notes master onto a new notes slide.
    # The header, date and footer stay on the master, as they do in PowerPoint.
    CLONEABLE = %i[SLIDE_IMAGE BODY SLIDE_NUMBER].freeze

    def shapes
      @shapes ||= NotesSlideShapes.new(@element.spTree, self)
    end

    def placeholders
      @placeholders ||= NotesSlidePlaceholders.new(@element.spTree, self)
    end

    # The body placeholder that holds the notes text, or nil if it was removed.
    def notes_placeholder
      placeholders.find { |ph| ph.placeholder_format.type.name == :BODY }
    end

    # The text frame of the notes placeholder, or nil when there is none.
    #
    # @return [TextFrame, nil]
    def notes_text_frame
      notes_placeholder&.text_frame
    end

    # Copy the cloneable placeholders from +notes_master+, keeping their order.
    # Called once, when the notes slide is created.
    def clone_master_placeholders(notes_master)
      notes_master.shapes.each do |shape|
        next unless shape.placeholder?
        next unless CLONEABLE.include?(shape.element.ph_type.name)

        shapes.clone_placeholder(shape)
      end
      self
    end

    def inspect
      "#<Pptx::NotesSlide #{part.partname}>"
    end
  end

  # One slide in a presentation.
  class Slide < BaseSlide
    include PatternMatching

    pattern_keys :slide_id, :name, :layout, :shapes, :placeholders

    # The id that identifies this slide within the presentation, stable across
    # reordering.
    def slide_id
      part.slide_id
    end

    # The layout this slide takes its appearance from.
    def layout
      part.slide_layout
    end

    # True when this slide inherits the master's background.
    def follows_master_background?
      @element.bg.nil?
    end

    def notes_slide?
      part.notes_slide?
    end

    # This slide's speaker notes page, created on first use.
    #
    # Use {#notes_slide?} to ask whether there is one without creating it, or
    # {#notes} and {#notes=} for the common case of just the text.
    #
    # @return [NotesSlide]
    def notes_slide
      part.notes_slide
    end

    # The speaker notes as plain text, or nil when the slide has none.
    #
    # Reading never creates a notes slide; that would add parts to a file
    # merely by looking at it.
    def notes
      notes_slide? ? notes_slide.notes_text_frame&.text : nil
    end

    # Replace the speaker notes, creating the notes slide if need be.
    def notes=(text)
      frame = notes_slide.notes_text_frame
      raise Error, "this slide's notes page has no notes placeholder" if frame.nil?

      frame.text = text
    end

    # The shapes on this slide, in z-order.
    def shapes
      @shapes ||= SlideShapes.new(@element.spTree, self)
    end

    # The placeholders on this slide, keyed by `idx`.
    def placeholders
      @placeholders ||= SlidePlaceholders.new(@element.spTree, self)
    end

    def inspect
      "#<Pptx::Slide id=#{slide_id} #{part.partname}>"
    end
  end

  # The slides of a presentation, in order.
  #
  # Enumerable, and indexable by position.
  class Slides < ParentedElementProxy
    include DeconstructToArray

    include Enumerable

    def initialize(sld_id_list, presentation)
      super
      @sld_id_list = sld_id_list
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @sld_id_list.sldId_list.each { |sld_id| yield part.related_slide(sld_id.rId) }
      self
    end

    def size
      @sld_id_list.size
    end
    alias length size

    def empty?
      size.zero?
    end

    # Indexed access, supporting a negative index as Ruby arrays do.
    #
    # @return [Slide, nil] nil when +index+ is out of range
    def [](index, length = nil)
      slice_members(@sld_id_list.sldId_list, index, length) { |entry| part.related_slide(entry.rId) }
    end

    # As {#[]}, but raises rather than returning nil.
    def fetch(index)
      self[index] || raise(IndexError, "slide index #{index} out of range")
    end

    # Enumerable gives us #first but not #last.
    def last
      self[-1]
    end

    # The slide with the given slide id.
    #
    # @return [Slide, nil]
    def by_id(slide_id)
      part.slide_by_id(slide_id)
    end

    # Add a slide inheriting from +slide_layout+, and return it.
    #
    # The layout's placeholders are copied onto the new slide, preserving
    # z-order. Date, footer and slide number come along only with
    # +footers: true+, which is what PowerPoint needs to show them.
    def add(slide_layout, footers: false)
      r_id, slide = part.add_slide(slide_layout)
      slide.shapes.clone_layout_placeholders(slide_layout, footers: footers)
      @sld_id_list.add_slide_id(r_id)
      slide
    end

    # The zero-based position of +slide+.
    #
    # @return [Integer, nil] nil when the slide is not in this collection
    def index(slide)
      each_with_index.find { |s, _| s == slide }&.last
    end

    def inspect
      "#<Pptx::Slides size=#{size}>"
    end
  end

  # A slide layout: the arrangement a slide inherits from.
  class SlideLayout < BaseSlide
    include PatternMatching

    pattern_keys :name, :type, :slide_master, :shapes, :placeholders

    # Placeholders a new slide leaves out unless asked: PowerPoint shows them
    # only where the slide carries them, which is its Header & Footer switch.
    LATENT_PLACEHOLDER_TYPES = [
      Enum::PP_PLACEHOLDER::DATE,
      Enum::PP_PLACEHOLDER::FOOTER,
      Enum::PP_PLACEHOLDER::SLIDE_NUMBER
    ].freeze

    # The shapes on this layout.
    def shapes
      @shapes ||= LayoutShapes.new(@element.spTree, self)
    end

    # The placeholders on this layout, in `idx` order.
    def placeholders
      @placeholders ||= LayoutPlaceholders.new(@element.spTree, self)
    end

    # The placeholders a new slide based on this layout should receive.
    def cloneable_placeholders
      placeholders.reject { |ph| LATENT_PLACEHOLDER_TYPES.include?(ph.element.ph_type) }
    end

    # The master this layout inherits from.
    def slide_master
      part.slide_master
    end

    # The kind of layout this is, e.g. "title" or "obj". PowerPoint uses it to
    # decide which layout to offer for a given command.
    def type
      @element.type
    end

    def type=(value)
      @element.type = value&.to_s
    end

    # The slides based on this layout.
    def used_by_slides
      part.package.presentation_part.presentation.slides.select { |s| s.layout == self }
    end

    def inspect
      "#<Pptx::SlideLayout #{name.inspect}>"
    end
  end

  # The layouts belonging to one slide master.
  class SlideLayouts < ParentedElementProxy
    include DeconstructToArray

    include Enumerable

    def initialize(sld_layout_id_list, slide_master)
      super
      @sld_layout_id_list = sld_layout_id_list
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @sld_layout_id_list.sldLayoutId_list.each do |entry|
        yield part.related_slide_layout(entry.rId)
      end
      self
    end

    def size
      @sld_layout_id_list.size
    end
    alias length size

    # Indexed by position, or looked up by layout name.
    #
    #   layouts[1]
    #   layouts["Title and Content"]
    def [](key, length = nil)
      return by_name(key) if key.is_a?(String)

      slice_members(@sld_layout_id_list.sldLayoutId_list, key, length) do |entry|
        part.related_slide_layout(entry.rId)
      end
    end

    def fetch(key)
      self[key] || raise(IndexError, "no slide layout #{key.inspect}")
    end

    # @return [SlideLayout, nil]
    def by_name(name)
      find { |layout| layout.name == name }
    end

    # @return [Integer, nil]
    def index(slide_layout)
      each_with_index.find { |l, _| l == slide_layout }&.last
    end

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

    # Add a layout to this master.
    #
    #   layout = master.slide_layouts.add("Section Header", type: "secHead")
    #   layout.placeholders.add(:title, at: [x, y], size: [cx, cy])
    #
    # The layout starts empty; placeholders are added to it explicitly, which
    # is what decides what a slide using it inherits.
    #
    # @param name [String] the name PowerPoint shows in the layout gallery
    # @param type [String, Symbol, nil] one of the 36 ST_SlideLayoutType values
    # @return [SlideLayout]
    def add(name, type: "obj")
      master_part = parent.part
      layout_part = Parts::SlideLayoutPart.new_layout(master_part.package, master_part, name,
                                                      type&.to_s)
      master_part.add_slide_layout(layout_part)
      layout = layout_part.slide_layout
      yield layout if block_given?
      layout
    end

    def inspect
      "#<Pptx::SlideLayouts size=#{size}>"
    end
  end

  # A slide master.
  class SlideMaster < BaseMaster
    # The layouts that inherit from this master.
    def slide_layouts
      @slide_layouts ||= SlideLayouts.new(@element.get_or_add_sldLayoutIdLst, self)
    end

    # The colours and fonts this master draws from.
    #
    # @return [Pptx::Theme]
    def theme
      part.theme_part.theme
    end

    # The placeholders every layout and slide under this master inherits from.
    #
    # @return [MasterPlaceholders]
    def placeholders
      @placeholders ||= MasterPlaceholders.new(@element.spTree, self)
    end

    # The default text formatting per outline level for titles, body text and
    # everything else, which every layout and slide under this master inherits.
    #
    # @return [TextStyles]
    def text_styles
      TextStyles.new(@element.get_or_add_txStyles)
    end

    def inspect
      "#<Pptx::SlideMaster #{part.partname}>"
    end
  end

  # The slide masters of a presentation.
  class SlideMasters < ParentedElementProxy
    include DeconstructToArray

    include Enumerable

    def initialize(sld_master_id_list, presentation)
      super
      @sld_master_id_list = sld_master_id_list
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @sld_master_id_list.sldMasterId_list.each do |entry|
        yield part.related_slide_master(entry.rId)
      end
      self
    end

    def size
      @sld_master_id_list.size
    end
    alias length size

    def [](index, length = nil)
      slice_members(@sld_master_id_list.sldMasterId_list, index, length) do |entry|
        part.related_slide_master(entry.rId)
      end
    end

    # Add a slide master to the presentation, with its own theme.
    #
    #   master = deck.slide_masters.add(name: "Corporate")
    #   master.theme.colors.update(accent1: "1F497D")
    #   master.slide_layouts.add("Title Slide", type: "title")
    #
    # The master starts with the five placeholders PowerPoint expects of one --
    # title, body, date, footer and slide number -- positioned proportionally
    # to this presentation's slide size. Pass <tt>placeholders: :none</tt> to
    # get a bare master and place them yourself.
    #
    # @param name [String] names both the master's theme and its schemes
    # @param placeholders [Symbol] :standard or :none
    # @return [SlideMaster]
    def add(name: "Office Theme", placeholders: :standard)
      unless %i[standard none].include?(placeholders)
        raise ArgumentError, "placeholders must be :standard or :none, got #{placeholders.inspect}"
      end

      master_part = Parts::SlideMasterPart.new_master(part.package, name: name)
      register(master_part)
      master = master_part.slide_master
      master.placeholders.add_standard_set(slide_size) if placeholders == :standard
      yield master if block_given?
      master
    end

    def inspect
      "#<Pptx::SlideMasters size=#{size}>"
    end

    private

    # Master ids are numbered from 2147483648 upwards, above the slide-id range.
    def register(master_part)
      r_id = part.relate_to(master_part, Opc::RELATIONSHIP_TYPE::SLIDE_MASTER)
      entry = @sld_master_id_list.add_sldMasterId
      entry.rId = r_id
      used = @sld_master_id_list.sldMasterId_list.filter_map(&:id)
      entry.id = used.empty? ? 2_147_483_648 : used.max + 1
      r_id
    end

    def slide_size
      size_element = part.element.sldSz
      [size_element.cx, size_element.cy]
    end
  end
end
