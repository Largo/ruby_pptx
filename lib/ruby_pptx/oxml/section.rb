# frozen_string_literal: true

require "securerandom"
require "ruby_pptx/oxml/element"
require "ruby_pptx/oxml/content_model"
require "ruby_pptx/oxml/simple_types"

module Pptx
  module Oxml
    # `p:extLst`, the extension list on `p:presentation`.
    class CT_PresentationExtensionList < Element
      tag "p:extLst"
      zero_or_more "p:ext", as: :ext
    end

    # `p:ext`, one extension, identified by its URI.
    class CT_PresentationExtension < Element
      tag "p:ext"
      required_attr "uri", type: SimpleTypes::XsdString
    end

    # `p14:sldId`, a slide's membership in a section.
    #
    # Note this refers to the slide by its presentation-wide `p:sldId/@id`,
    # not by a relationship.
    class CT_SectionSlideId < Element
      tag "p14:sldId"
      required_attr "id", type: SimpleTypes::ST_SlideId
    end

    # `p14:sldIdLst`, the slides belonging to one section.
    class CT_SectionSlideIdList < Element
      tag "p14:sldIdLst"
      zero_or_more "p14:sldId", as: :sldId
    end

    # `p14:section`, one named section of a presentation.
    class CT_Section < Element
      tag "p14:section"
      required_attr "name", type: SimpleTypes::XsdString
      required_attr "id", type: SimpleTypes::XsdString
      one_and_only_one "p14:sldIdLst"

      # PowerPoint identifies a section by a brace-wrapped uppercase GUID.
      def self.new_id
        "{#{SecureRandom.uuid.upcase}}"
      end
    end

    # `p14:sectionLst`, the sections of a presentation.
    class CT_SectionList < Element
      tag "p14:sectionLst"
      zero_or_more "p14:section", as: :section

      def add_section(name, id = CT_Section.new_id)
        build_from_xml(
          %(<p14:section #{Ns.nsdecls("p14")} name="#{escape(name)}" id="#{id}">) \
          "<p14:sldIdLst/></p14:section>"
        ).tap { |element| append(element) }
      end

      def escape(text)
        text.to_s.gsub("&", "&amp;").gsub("<", "&lt;").gsub(">", "&gt;").gsub('"', "&quot;")
      end
    end
  end
end
