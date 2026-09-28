# frozen_string_literal: true

require "ruby_pptx/opc/package"
require "ruby_pptx/opc/constants"
require "ruby_pptx/oxml/presentation"
require "ruby_pptx/parts/slide"
require "ruby_pptx/presentation"

module Pptx
  module Parts
    # The presentation part, `/ppt/presentation.xml`; the package's main
    # document part.
    class PresentationPart < Opc::XmlPart
      def presentation = @presentation ||= Presentation.new(element, self)

      # Create a blank slide part inheriting from +slide_layout+.
      #
      # @return [Array(String, Pptx::Slide)] the new relationship id and slide
      def add_slide(slide_layout)
        slide_part = SlidePart.new_slide(next_slide_partname, package, slide_layout.part)
        r_id = relate_to(slide_part, Opc::RELATIONSHIP_TYPE::SLIDE)
        [r_id, slide_part.slide]
      end

      def core_properties = package.core_properties

      # The presentation's notes master, created from the default template on
      # first use.
      #
      # Like python-pptx, this relates the new part to the presentation but
      # does not add a `p:notesMasterIdLst` entry for it.
      def notes_master_part
        existing = rels.find { |r| r.reltype == Opc::RELATIONSHIP_TYPE::NOTES_MASTER }
        return existing.target_part if existing

        part = NotesMasterPart.create_default(package)
        relate_to(part, Opc::RELATIONSHIP_TYPE::NOTES_MASTER)
        part
      end

      def related_slide(r_id) = related_part(r_id).slide

      def related_slide_master(r_id) = related_part(r_id).slide_master

      # The slide with presentation-wide id +slide_id+.
      #
      # @return [Slide, nil]
      def slide_by_id(slide_id)
        entry = element.sldIdLst&.sldId_list&.find { |s| s.id == slide_id }
        entry && related_part(entry.rId).slide
      end

      # The slide id assigned to +slide_part+.
      def slide_id_for(slide_part)
        entry = element.sldIdLst&.sldId_list&.find { |s| related_part(s.rId).equal?(slide_part) }
        entry&.id || raise(NotFoundError, "slide part is not in this presentation")
      end

      # Renumber the slide parts so their partnames form a gapless sequence in
      # the order the slides appear.
      #
      # Slide order lives in `p:sldIdLst`, not in the partnames, so the two can
      # drift apart after slides are added or deleted. PowerPoint tolerates
      # that, but keeping them aligned makes saved packages predictable.
      def renumber_slide_parts(r_ids)
        r_ids.each_with_index do |r_id, index|
          related_part(r_id).partname = Opc::PackURI.new("/ppt/slides/slide#{index + 1}.xml")
        end
      end

      def next_slide_partname
        count = element.get_or_add_sldIdLst.size
        Opc::PackURI.new("/ppt/slides/slide#{count + 1}.xml")
      end

      def save(path_or_stream) = package.save(path_or_stream)
    end
  end
end
