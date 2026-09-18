# frozen_string_literal: true

require "time"
require "date"
require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/oxml/ns"

module Pptx
  module Oxml
    # `cp:coreProperties`, the root of `/docProps/core.xml`.
    #
    # Implements the Dublin Core metadata elements. Text properties read as an
    # empty string when absent and are capped at 255 characters, as the schema
    # requires.
    class CT_CoreProperties < Element
      tag "cp:coreProperties"

      TEXT_PROPERTIES = {
        category: "cp:category",
        content_status: "cp:contentStatus",
        author: "dc:creator",
        comments: "dc:description",
        identifier: "dc:identifier",
        keywords: "cp:keywords",
        language: "dc:language",
        last_modified_by: "cp:lastModifiedBy",
        subject: "dc:subject",
        title: "dc:title",
        version: "cp:version"
      }.freeze

      DATETIME_PROPERTIES = {
        created: "dcterms:created",
        last_printed: "cp:lastPrinted",
        modified: "dcterms:modified"
      }.freeze

      # `dcterms:created` and `dcterms:modified` must carry an explicit
      # xsi:type; the others must not.
      W3CDTF_TYPED = %i[created modified].freeze

      MAX_TEXT_LENGTH = 255

      (TEXT_PROPERTIES.values + DATETIME_PROPERTIES.values + ["cp:revision"]).each do |nsptag|
        zero_or_one nsptag, successors: []
      end

      def self.new_element
        Element.parse(%(<cp:coreProperties #{Ns.nsdecls("cp", "dc", "dcterms")}/>))
      end

      TEXT_PROPERTIES.each do |name, nsptag|
        define_method(name) { find(nsptag)&.text.to_s }

        define_method("#{name}=") do |value|
          value = value.to_s
          if value.length > MAX_TEXT_LENGTH
            raise ArgumentError,
                  "#{name} exceeds the #{MAX_TEXT_LENGTH}-character limit (got #{value.length})"
          end

          public_send("get_or_add_#{Ns.split_tag(nsptag).last}").text = value
        end
      end

      DATETIME_PROPERTIES.each do |name, nsptag|
        define_method(name) do
          raw = find(nsptag)&.text
          raw.nil? || raw.empty? ? nil : CT_CoreProperties.parse_w3cdtf(raw)
        end

        define_method("#{name}=") do |value|
          raise TypeError, "#{name} requires a Time, got #{value.class}" unless value.is_a?(Time)

          element = public_send("get_or_add_#{Ns.split_tag(nsptag).last}")
          element.text = value.getutc.strftime("%Y-%m-%dT%H:%M:%SZ")
          if W3CDTF_TYPED.include?(name)
            # PowerPoint declares xsi on the root, not on each child that uses
            # it, so declare it here first and let the child find it in scope.
            declare_namespace("xsi")
            element.set("xsi:type", "dcterms:W3CDTF")
          end
          value
        end
      end

      # A non-integer, negative or absent revision all read as 0.
      def revision
        raw = find("cp:revision")&.text
        return 0 if raw.nil?

        value = Integer(raw, 10, exception: false) || 0
        value.negative? ? 0 : value
      end

      def revision=(value)
        unless value.is_a?(Integer) && value >= 1
          raise ArgumentError, "revision requires a positive Integer, got #{value.inspect}"
        end

        get_or_add_revision.text = value.to_s
      end

      # Parse a W3CDTF timestamp, which may be a year, a year-month, a date,
      # or a full timestamp with a `Z` or numeric offset. Returns UTC, or nil
      # when the value does not parse.
      #
      # Each format must consume the *whole* string. Ruby's `strptime` is happy
      # to parse a prefix and discard the rest, so without that check "%Y"
      # would match "2003-12-31T10:14:55" and silently yield 2003-01-01.
      FORMATS = ["%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%Y-%m", "%Y"].freeze

      def self.parse_w3cdtf(value)
        parseable = value[0, 19]
        offset = value[19..].to_s

        parts = FORMATS.lazy.filter_map { |format| full_match(parseable, format) }.first
        return nil if parts.nil?

        time = Time.utc(parts[:year], parts[:mon] || 1, parts[:mday] || 1,
                        parts[:hour] || 0, parts[:min] || 0, parts[:sec] || 0)
        apply_offset(time, offset)
      end

      # Parsed fields for +value+, but only if +format+ consumed all of it.
      def self.full_match(value, format)
        parts = Date._strptime(value, format)
        parts if parts && parts[:leftover].nil? && parts[:year]
      end
      private_class_method :full_match

      # A trailing "+02:00" means the stamp is local to that offset; shift it
      # back to UTC. A "Z" or an absent offset is already UTC.
      def self.apply_offset(time, offset)
        match = /\A([+-])(\d\d):(\d\d)\z/.match(offset)
        return time unless match

        sign = match[1] == "+" ? -1 : 1
        time + (sign * ((Integer(match[2], 10) * 3600) + (Integer(match[3], 10) * 60)))
      end
      private_class_method :apply_offset
    end
  end
end
