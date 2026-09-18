# frozen_string_literal: true

require "pptx/opc/package"
require "pptx/opc/constants"
require "pptx/oxml/slide"
require "pptx/slide"
require "pptx/parts/chart"

module Pptx
  module Parts
    # Common to every slide-like part: slides, layouts, masters, notes slides
    # and the notes master.
    class BaseSlidePart < Opc::XmlPart
      # The internal name of this slide.
      def name = element.cSld.name

      # Create a chart part for +chart_data+ and relate this slide to it.
      #
      # @return [String] the relationship id of the new chart part
      def add_chart_part(chart_type, chart_data)
        chart_part = ChartPart.new_chart(chart_type, chart_data, package)
        relate_to(chart_part, Opc::RELATIONSHIP_TYPE::CHART)
      end

      # Relate this slide to the media part holding +video+.
      #
      # Two relationships are made to the same part, one MEDIA and one VIDEO.
      # PowerPoint has embedded media two different ways over the years and
      # writes both so either era can find it.
      #
      # @return [Array(String, String)] the media rId and the video rId
      def get_or_add_video_media_part(video)
        media_part = package.get_or_add_media_part(video)
        [relate_to(media_part, Opc::RELATIONSHIP_TYPE::MEDIA),
         relate_to(media_part, Opc::RELATIONSHIP_TYPE::VIDEO)]
      end

      # The image part for +image_file+, related to this slide.
      #
      # Both the part and the relationship are reused when they already exist,
      # so the same image added twice is stored once.
      #
      # @return [Array(Pptx::Parts::ImagePart, String)] the part and its rId
      def get_or_add_image_part(image_file)
        image_part = package.get_or_add_image_part(image_file)
        [image_part, relate_to(image_part, Opc::RELATIONSHIP_TYPE::IMAGE)]
      end
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
