# frozen_string_literal: true

require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/oxml/simple_types"
require "pptx/oxml/section"
require "pptx/oxml/ns"

module Pptx
  module Oxml
    ST = SimpleTypes

    # `p:presentation`, the root of `/ppt/presentation.xml`.
    class CT_Presentation < Element
      tag "p:presentation"
      zero_or_one "p:sldMasterIdLst",
                  successors: %w[p:notesMasterIdLst p:handoutMasterIdLst p:sldIdLst p:sldSz p:notesSz]
      zero_or_one "p:sldIdLst", successors: %w[p:sldSz p:notesSz]
      zero_or_one "p:sldSz", successors: %w[p:notesSz]
      zero_or_one "p:extLst", successors: []

      # The URI PowerPoint uses to mark the section-list extension.
      SECTION_LIST_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"

      # The `p14:sectionLst` element, or nil when the presentation has no
      # sections.
      def section_list
        xpath("./p:extLst/p:ext/p14:sectionLst").first
      end

      # The `p14:sectionLst`, created inside its extension if absent.
      def get_or_add_section_list
        section_list || begin
          ext = get_or_add_extLst.add_ext(uri: SECTION_LIST_EXT_URI)
          ext.append(ext.build_from_xml(%(<p14:sectionLst #{Ns.nsdecls("p14")}/>)))
          section_list
        end
      end

      def remove_section_list
        extension = xpath("./p:extLst/p:ext").find { |e| e.get("uri") == SECTION_LIST_EXT_URI }
        return if extension.nil?

        extLst.remove(extension)
        remove_extLst if extLst.ext_list.empty?
      end
    end

    # `p:sldId`, a reference from the presentation to one slide.
    class CT_SlideId < Element
      tag "p:sldId"
      required_attr "id", type: ST::ST_SlideId
      required_attr "r:id", type: ST::XsdString, as: :rId
    end

    # `p:sldIdLst`, the ordered list of slides in the presentation.
    class CT_SlideIdList < Element
      tag "p:sldIdLst"
      zero_or_more "p:sldId", as: :sldId

      MIN_SLIDE_ID = 256
      MAX_SLIDE_ID = 2_147_483_647

      def add_slide_id(r_id) = add_sldId(id: next_id, rId: r_id)

      def size = sldId_list.size

      private

      # One greater than the highest id in use, which avoids reusing the id of
      # a deleted slide. Only if that would overflow do we hunt for a gap.
      def next_id
        used = sldId_list.map(&:id)
        simple_next = ([MIN_SLIDE_ID - 1] + used).max + 1
        return simple_next if simple_next <= MAX_SLIDE_ID

        valid = used.grep(MIN_SLIDE_ID..MAX_SLIDE_ID).sort
        return MIN_SLIDE_ID if valid.empty?

        valid.each_with_index do |used_id, i|
          candidate = MIN_SLIDE_ID + i
          return candidate if candidate != used_id
        end
        MIN_SLIDE_ID + valid.size
      end
    end

    # `p:sldMasterIdLst`, the slide masters belonging to the presentation.
    class CT_SlideMasterIdList < Element
      tag "p:sldMasterIdLst"
      zero_or_more "p:sldMasterId", as: :sldMasterId

      def size = sldMasterId_list.size
    end

    # `p:sldMasterId`, a reference to one slide master.
    class CT_SlideMasterIdListEntry < Element
      tag "p:sldMasterId"
      required_attr "r:id", type: ST::XsdString, as: :rId
      optional_attr "id", type: ST::ST_SlideMasterId
    end

    # `p:sldSz`, the slide dimensions for the whole presentation.
    class CT_SlideSize < Element
      tag "p:sldSz"
      required_attr "cx", type: ST::ST_SlideSizeCoordinate
      required_attr "cy", type: ST::ST_SlideSizeCoordinate
    end
  end
end
