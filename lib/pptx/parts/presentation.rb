# frozen_string_literal: true

require "pptx/opc/package"
require "pptx/opc/constants"
require "pptx/oxml/presentation"
require "pptx/parts/slide"
require "pptx/presentation"

module Pptx
  module Parts
    # The presentation part, `/ppt/presentation.xml`; the package's main
    # document part.
    class PresentationPart < Opc::XmlPart
      def presentation = @presentation ||= Presentation.new(element, self)

      def core_properties = package.core_properties

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
