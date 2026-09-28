# frozen_string_literal: true

require "ruby_pptx/pattern_matching"

require "ruby_pptx/element_proxy"
require "ruby_pptx/slide"
require "ruby_pptx/section"

module Pptx
  # A PowerPoint presentation.
  #
  # Open one with {Presentation.open}, or start from the built-in template with
  # {Presentation.new}:
  #
  #   prs = Pptx::Presentation.open("deck.pptx")
  #   prs.slides.each { |slide| puts slide.name }
  #   prs.slide_width = Pptx.inches(13.333)
  #   prs.save("wide.pptx")
  class Presentation < PartElementProxy
    include PatternMatching

    pattern_keys :slides, :slide_layouts, :slide_masters, :slide_width, :slide_height

    class << self
      # Open a presentation from a path or an IO stream.
      #
      # Given a block, yields the presentation and returns the block's value,
      # as `File.open` does.
      #
      # @raise [PackageNotFoundError] when the file is not an OPC package
      # @raise [Error] when the package is not a PowerPoint presentation
      def open(pptx = nil)
        pptx ||= default_template_path
        presentation_part = Package.open(pptx).main_document_part
        unless VALID_CONTENT_TYPES.include?(presentation_part.content_type)
          raise Error,
                "not a PowerPoint file; content type is #{presentation_part.content_type.inspect}"
        end

        presentation = presentation_part.presentation
        block_given? ? yield(presentation) : presentation
      end

      # A new presentation based on the built-in default template.
      def new_default
        Presentation.open(nil)
      end

      def default_template_path
        File.expand_path("templates/default.pptx", __dir__)
      end
    end

    VALID_CONTENT_TYPES = [
      Opc::CONTENT_TYPE::PML_PRESENTATION_MAIN,
      Opc::CONTENT_TYPE::PML_PRES_MACRO_MAIN
    ].freeze

    # The Dublin Core metadata for this presentation.
    def core_properties
      part.core_properties
    end

    # The slides in this presentation.
    def slides
      @slides ||= begin
        sld_id_list = @element.get_or_add_sldIdLst
        part.renumber_slide_parts(sld_id_list.sldId_list.map(&:rId))
        Slides.new(sld_id_list, self)
      end
    end

    # The slide masters in this presentation.
    def slide_masters
      @slide_masters ||= SlideMasters.new(@element.get_or_add_sldMasterIdLst, self)
    end

    # The notes master, created from the default template on first use.
    #
    # @return [NotesMaster]
    def notes_master
      part.notes_master_part.notes_master
    end

    # The first slide master, which is the only one in most presentations.
    def slide_master
      slide_masters[0]
    end

    # The layouts of the first slide master.
    #
    # A presentation may have several masters, each with its own layouts; this
    # is a convenience for the common case of one.
    def slide_layouts
      slide_master.slide_layouts
    end

    # The sections grouping this presentation's slides.
    #
    # Sections are a PowerPoint 2010 extension; a presentation with none has
    # an empty collection and writes no extension element.
    def sections
      @sections ||= Sections.new(self)
    end

    # @return [Length, nil] nil when the presentation defines no slide size
    def slide_width
      @element.sldSz&.cx
    end

    def slide_width=(width)
      @element.get_or_add_sldSz.cx = width
    end

    # @return [Length, nil]
    def slide_height
      @element.sldSz&.cy
    end

    def slide_height=(height)
      @element.get_or_add_sldSz.cy = height
    end

    # Write this presentation to a path or an IO stream.
    def save(path_or_stream)
      part.save(path_or_stream)
      self
    end

    # The .pptx bytes, for callers who want to hand them to something other
    # than the filesystem.
    #
    # @return [String]
    def to_blob
      buffer = StringIO.new(+"", "w+b")
      save(buffer)
      buffer.string
    end

    def inspect
      "#<Pptx::Presentation slides=#{slides.size} #{slide_width&.inches}x#{slide_height&.inches}in>"
    end
  end
end
