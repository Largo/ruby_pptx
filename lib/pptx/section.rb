# frozen_string_literal: true

require "pptx/pattern_matching"

require "pptx/sliceable"

require "pptx/element_proxy"

module Pptx
  # The sections of a presentation.
  #
  # Sections group slides in the thumbnail pane and in presenter view. They are
  # a PowerPoint 2010 extension rather than part of the base PresentationML
  # schema, and python-pptx does not model them at all.
  #
  #   prs.sections.add("Introduction", slides: prs.slides.first(2))
  #   prs.sections["Introduction"].slides.map(&:slide_id)
  class Sections
    include DeconstructToArray

    include Sliceable
    include Enumerable

    def initialize(presentation)
      @presentation = presentation
    end

    def each
      return enum_for(:each) { size } unless block_given?

      section_elements.each { |element| yield Section.new(element, @presentation) }
      self
    end

    def size = section_elements.size
    alias length size

    def empty? = size.zero?

    # Indexed by position, or looked up by name.
    def [](key, length = nil)
      return by_name(key) if key.is_a?(String)

      slice_members(section_elements, key, length) { |element| Section.new(element, @presentation) }
    end

    # @return [Section, nil]
    def by_name(name) = find { |section| section.name == name }

    # Add a section, optionally containing +slides+.
    #
    # @return [Section]
    def add(name, slides: [])
      element = @presentation.element.get_or_add_section_list.add_section(name)
      Section.new(element, @presentation).tap do |section|
        slides.each { |slide| section << slide }
      end
    end

    # Remove a section. Its slides stay in the presentation; only the grouping
    # goes away, which is what PowerPoint does when a section is removed
    # without its slides.
    def delete(section)
      list = @presentation.element.section_list
      return nil if list.nil?

      list.remove(section.element)
      @presentation.element.remove_section_list if list.section_list.empty?
      section
    end

    def inspect = "#<Pptx::Sections #{map(&:name).inspect}>"

    private

    def section_elements = @presentation.element.section_list&.section_list || []
  end

  # One named section of a presentation.
  class Section < ElementProxy
    include PatternMatching

    pattern_keys :name, :id, :slides

    def initialize(element, presentation)
      super(element)
      @presentation = presentation
    end

    def name = @element.name

    def name=(value)
      @element.name = value.to_s
    end

    # The GUID PowerPoint uses to identify this section.
    def id = @element.id

    # The slides in this section, in section order.
    #
    # A section stores slide ids, so a slide deleted from the presentation
    # simply drops out here rather than leaving a dangling entry.
    def slides
      slide_ids.filter_map { |slide_id| @presentation.slides.by_id(slide_id) }
    end

    def slide_ids = @element.sldIdLst.sldId_list.map(&:id)

    # Put +slide+ in this section. A slide belongs to at most one section, so
    # it is removed from any other first.
    def <<(slide)
      @presentation.sections.each { |other| other.delete_slide(slide) unless other == self }
      return self if slide_ids.include?(slide.slide_id)

      @element.sldIdLst.add_sldId(id: slide.slide_id)
      self
    end
    alias add_slide <<

    def delete_slide(slide)
      entry = @element.sldIdLst.sldId_list.find { |e| e.id == slide.slide_id }
      @element.sldIdLst.remove(entry) if entry
      slide
    end

    def include?(slide) = slide_ids.include?(slide.slide_id)

    def inspect = "#<Pptx::Section #{name.inspect} slides=#{slide_ids.size}>"
  end
end
