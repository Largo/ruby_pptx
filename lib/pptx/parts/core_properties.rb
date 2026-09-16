# frozen_string_literal: true

require "pptx/opc/package"
require "pptx/opc/constants"
require "pptx/oxml/core_properties"

module Pptx
  module Parts
    # `/docProps/core.xml`, the Dublin Core metadata for the package.
    #
    # Reads and writes are delegated straight to the root element, so this
    # class is mostly a named home for them.
    class CorePropertiesPart < Opc::XmlPart
      DELEGATED = (Oxml::CT_CoreProperties::TEXT_PROPERTIES.keys +
                   Oxml::CT_CoreProperties::DATETIME_PROPERTIES.keys +
                   [:revision]).freeze

      DELEGATED.each do |name|
        define_method(name) { element.public_send(name) }
        define_method("#{name}=") { |value| element.public_send("#{name}=", value) }
      end

      # A core-properties part for a package that has none, matching what
      # PowerPoint would write for a new presentation.
      def self.default(package)
        part = new(
          Opc::PackURI.new("/docProps/core.xml"),
          Opc::CONTENT_TYPE::OPC_CORE_PROPERTIES,
          package,
          Oxml::CT_CoreProperties.new_element
        )
        part.title = "PowerPoint Presentation"
        part.last_modified_by = "ruby_pptx"
        part.revision = 1
        part.modified = Time.now.utc
        part
      end
    end
  end
end
