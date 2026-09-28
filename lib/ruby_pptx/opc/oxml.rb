# frozen_string_literal: true

require "ruby_pptx/oxml/element"
require "ruby_pptx/oxml/content_model"
require "ruby_pptx/oxml/simple_types"
require "ruby_pptx/opc/constants"

module Pptx
  module Opc
    # Element classes for the two XML formats the packaging layer owns:
    # `[Content_Types].xml` and the `.rels` items.
    module Oxml
      ST = Pptx::Oxml::SimpleTypes
      Element = Pptx::Oxml::Element

      # Serialize a part's root element as the bytes of a `.xml` file,
      # including the XML declaration Office expects.
      #
      # The declaration is written with single quotes and `standalone="yes"`
      # to match what python-pptx (and lxml) produce, so packages compare
      # byte-for-byte.
      def self.serialize_part_xml(element)
        node = element.is_a?(Element) ? element.node : element
        body = node.to_xml(save_with: Nokogiri::XML::Node::SaveOptions::AS_XML |
                                      Nokogiri::XML::Node::SaveOptions::NO_DECLARATION)
        %(<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n#{body})
      end

      # `<Default>`: the content type for every part with a given extension.
      class CT_Default < Element
        tag "ct:Default"
        required_attr "Extension", type: ST::ST_Extension, as: :extension
        required_attr "ContentType", type: ST::ST_ContentType, as: :contentType
      end

      # `<Override>`: the content type for one specific part.
      class CT_Override < Element
        tag "ct:Override"
        required_attr "PartName", type: ST::XsdAnyUri, as: :partName
        required_attr "ContentType", type: ST::ST_ContentType, as: :contentType
      end

      # `<Types>`, the root of `[Content_Types].xml`.
      class CT_Types < Element
        tag "ct:Types"
        zero_or_more "ct:Default", as: :default
        zero_or_more "ct:Override", as: :override

        def self.new_element
          Element.parse(%(<Types xmlns="#{NAMESPACE::OPC_CONTENT_TYPES}"/>))
        end

        # Named `*_for` rather than `add_default` / `add_override`: those names
        # belong to the accessors the content-model macros generate.
        def add_default_for(ext, content_type)
          add_default(extension: ext, contentType: content_type)
        end

        def add_override_for(partname, content_type)
          add_override(partName: partname.to_s, contentType: content_type)
        end
      end

      # `<Relationship>`: one link from a part (or the package) to a target.
      class CT_Relationship < Element
        tag "pr:Relationship"
        required_attr "Id", type: ST::XsdId, as: :rId
        required_attr "Type", type: ST::XsdAnyUri, as: :reltype
        required_attr "Target", type: ST::XsdAnyUri, as: :target_ref
        optional_attr "TargetMode", type: ST::ST_TargetMode,
                                    default: RELATIONSHIP_TARGET_MODE::INTERNAL, as: :targetMode

        def external?
          targetMode == RELATIONSHIP_TARGET_MODE::EXTERNAL
        end
      end

      # `<Relationships>`, the root of a `.rels` item.
      class CT_Relationships < Element
        tag "pr:Relationships"
        zero_or_more "pr:Relationship", as: :relationship

        def self.new_element
          Element.parse(%(<Relationships xmlns="#{NAMESPACE::OPC_RELATIONSHIPS}"/>))
        end

        def add_rel(r_id, reltype, target, is_external: false)
          target_mode = if is_external
                          RELATIONSHIP_TARGET_MODE::EXTERNAL
                        else
                          RELATIONSHIP_TARGET_MODE::INTERNAL
                        end
          add_relationship(rId: r_id, reltype: reltype, target_ref: target, targetMode: target_mode)
        end

        # Bytes for a `.rels` item, with the XML declaration.
        def xml_file_bytes
          Oxml.serialize_part_xml(self)
        end
      end
    end
  end
end
