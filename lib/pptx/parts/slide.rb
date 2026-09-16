# frozen_string_literal: true

require "pptx/opc/package"
require "pptx/opc/constants"
require "pptx/oxml/slide"
require "pptx/slide"

module Pptx
  module Parts
    # Common to every slide-like part: slides, layouts, masters, notes slides
    # and the notes master.
    class BaseSlidePart < Opc::XmlPart
      # The internal name of this slide.
      def name = element.cSld.name
    end

    # A slide part, `/ppt/slides/slideN.xml`.
    class SlidePart < BaseSlidePart
      # A new, empty slide part inheriting from +slide_layout_part+.
      def self.new_slide(partname, package, slide_layout_part)
        part = new(partname, Opc::CONTENT_TYPE::PML_SLIDE, package, Oxml::CT_Slide.new_element)
        part.relate_to(slide_layout_part, Opc::RELATIONSHIP_TYPE::SLIDE_LAYOUT)
        part
      end

      def slide = @slide ||= Slide.new(element, self)

      # The layout part this slide inherits from.
      def slide_layout
        part_related_by(Opc::RELATIONSHIP_TYPE::SLIDE_LAYOUT).slide_layout
      end

      # The presentation-wide id of this slide.
      def slide_id = package.presentation_part.slide_id_for(self)

      def notes_slide? = rels.any? { |r| r.reltype == Opc::RELATIONSHIP_TYPE::NOTES_SLIDE }
    end

    # A slide-layout part, `/ppt/slideLayouts/slideLayoutN.xml`.
    class SlideLayoutPart < BaseSlidePart
      def slide_layout = @slide_layout ||= SlideLayout.new(element, self)

      def slide_master
        part_related_by(Opc::RELATIONSHIP_TYPE::SLIDE_MASTER).slide_master
      end
    end

    # A slide-master part, `/ppt/slideMasters/slideMasterN.xml`.
    class SlideMasterPart < BaseSlidePart
      def slide_master = @slide_master ||= SlideMaster.new(element, self)

      def related_slide_layout(r_id) = related_part(r_id).slide_layout
    end

    # A notes-master part.
    class NotesMasterPart < BaseSlidePart
    end

    # A notes-slide part.
    class NotesSlidePart < BaseSlidePart
    end
  end
end
