# frozen_string_literal: true

require "nokogiri"

module Pptx
  module Spec
    # Validates generated parts against the published ISO/IEC 29500-4 schemas.
    #
    # This is the only oracle for a slide master: python-pptx can read one but
    # cannot create one, so there is nothing to diff a generated master
    # against. The schema is the actual standard, which is a stronger check
    # than comparing against another implementation's output.
    #
    # The schemas are not redistributed with this gem. Point OOXML_SCHEMAS at a
    # directory holding `pml.xsd`, `dml-main.xsd` and the `shared-*.xsd` files
    # they import, and these examples run; without it they skip. See PORTING.md.
    module Schema
      # Parts are validated against the schema that defines their root element.
      PML_ROOTS = %w[sld sldLayout sldMaster notesMaster notes presentation
                     presentationPr viewPr tableStyles].freeze

      class << self
        def dir = ENV["OOXML_SCHEMAS"]

        def available? = !dir.nil? && File.exist?(File.join(dir.to_s, "pml.xsd"))

        # Compiled once: loading dml-main.xsd and its imports is slow enough to
        # matter across a few dozen parts.
        def schema(name)
          @schemas ||= {}
          @schemas[name] ||= Dir.chdir(dir) { Nokogiri::XML::Schema(File.read("#{name}.xsd")) }
        end

        # @return [Array<String>] one message per violation, empty when valid
        def errors_for(xml, root_name)
          schema_name = PML_ROOTS.include?(root_name) ? "pml" : "dml-main"
          schema(schema_name).validate(Nokogiri::XML(xml)).map(&:message)
        end
      end

      # Validate every XML part of a saved package.
      #
      # @return [Hash{String => Array<String>}] part name => violations
      def schema_violations(path)
        violations = {}
        Zip::File.open(path) do |zip|
          zip.each do |entry|
            next unless entry.name.end_with?(".xml")
            next if entry.name.start_with?("_rels", "docProps") || entry.name.include?("/_rels/")
            next if entry.name == "[Content_Types].xml"

            xml = entry.get_input_stream.read
            root = Nokogiri::XML(xml).root
            errors = Schema.errors_for(xml, root.name)
            violations[entry.name] = errors unless errors.empty?
          end
        end
        violations
      end

      def require_schemas!
        return if Schema.available?

        raise "OOXML_SCHEMAS is not set and REQUIRE_SCHEMAS is" if ENV["REQUIRE_SCHEMAS"]

        skip "OOXML schemas not available; set OOXML_SCHEMAS to enable"
      end
    end
  end
end

RSpec.configure { |config| config.include Pptx::Spec::Schema }
