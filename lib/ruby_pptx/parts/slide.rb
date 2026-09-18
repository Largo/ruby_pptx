# frozen_string_literal: true

require "ruby_pptx/opc/package"
require "ruby_pptx/opc/constants"
require "ruby_pptx/oxml/slide"
require "ruby_pptx/slide"
require "ruby_pptx/parts/chart"

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

      # Create a combo chart part and relate this slide to it.
      #
      # @return [String] the relationship id of the new chart part
      def add_combo_chart_part(builder)
        chart_part = ChartPart.new_combo_chart(builder, package)
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
      PARTNAME_TEMPLATE = "/ppt/slideLayouts/slideLayout%d.xml"

      # A new layout named +name+, inheriting from +slide_master_part+.
      def self.new_layout(package, slide_master_part, name, layout_type)
        part = new(package.next_partname(PARTNAME_TEMPLATE),
                   Opc::CONTENT_TYPE::PML_SLIDE_LAYOUT, package,
                   Oxml::CT_SlideLayout.new_element(name, layout_type))
        part.relate_to(slide_master_part, Opc::RELATIONSHIP_TYPE::SLIDE_MASTER)
        part
      end

      def slide_layout = @slide_layout ||= SlideLayout.new(element, self)

      def slide_master
        part_related_by(Opc::RELATIONSHIP_TYPE::SLIDE_MASTER).slide_master
      end
    end

    # A slide-master part, `/ppt/slideMasters/slideMasterN.xml`.
    class SlideMasterPart < BaseSlidePart
      PARTNAME_TEMPLATE = "/ppt/slideMasters/slideMaster%d.xml"
      THEME_PARTNAME_TEMPLATE = "/ppt/theme/theme%d.xml"

      # A new master with its own theme part and no layouts yet.
      #
      # The theme is a separate part related from the master, which is where
      # the master's colours and fonts actually live.
      def self.new_master(package, name: "Office Theme")
        part = new(package.next_partname(PARTNAME_TEMPLATE),
                   Opc::CONTENT_TYPE::PML_SLIDE_MASTER, package,
                   Oxml::CT_SlideMaster.new_element)
        theme = ThemePart.new_theme(package.next_partname(THEME_PARTNAME_TEMPLATE), package,
                                    name: name)
        part.relate_to(theme, Opc::RELATIONSHIP_TYPE::THEME)
        part
      end

      def slide_master = @slide_master ||= SlideMaster.new(element, self)

      # The theme this master draws its colours and fonts from.
      def theme_part = part_related_by(Opc::RELATIONSHIP_TYPE::THEME)

      # Add +slide_layout_part+ to this master's layout list.
      #
      # @return [String] the relationship id
      def add_slide_layout(slide_layout_part)
        r_id = relate_to(slide_layout_part, Opc::RELATIONSHIP_TYPE::SLIDE_LAYOUT)
        entry = element.get_or_add_sldLayoutIdLst.add_sldLayoutId
        entry.rId = r_id
        entry.id = next_slide_layout_id
        r_id
      end

      def related_slide_layout(r_id) = related_part(r_id).slide_layout

      private

      # Layout ids are numbered from 2147483649 upwards, continuing from the
      # highest already in use so a removed layout does not free its id for
      # reuse -- PowerPoint keeps references to them elsewhere in the package.
      def next_slide_layout_id
        used = element.get_or_add_sldLayoutIdLst.sldLayoutId_list.filter_map(&:id)
        used.empty? ? 2_147_483_649 : used.max + 1
      end
    end

    # A notes-master part.
    class NotesMasterPart < BaseSlidePart
    end

    # A notes-slide part.
    class NotesSlidePart < BaseSlidePart
    end
  end
end
