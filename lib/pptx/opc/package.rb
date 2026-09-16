# frozen_string_literal: true

require "pptx/errors"
require "pptx/oxml/element"
require "pptx/opc/constants"
require "pptx/opc/oxml"
require "pptx/opc/pack_uri"
require "pptx/opc/serialized"

module Pptx
  module Opc
    # Relationship behaviour shared by the package and every part.
    module Relatable
      # The single part related to this one by +reltype+.
      #
      # @raise [NotFoundError] when there is no such relationship
      # @raise [Error] when there is more than one
      def part_related_by(reltype) = rels.part_with_reltype(reltype)

      # Relate this part to +target+, returning the rId. An existing matching
      # relationship is reused rather than duplicated.
      #
      # @param target [Part, String] a part, or a URL when +external:+
      def relate_to(target, reltype, external: false)
        if external
          rels.get_or_add_external(reltype, target)
        else
          rels.get_or_add(reltype, target)
        end
      end

      def related_part(r_id) = rels.fetch(r_id).target_part

      def target_ref(r_id) = rels.fetch(r_id).target_ref

      # @return [Relationships]
      def rels = raise(NotImplementedError, "#{self.class} must implement #rels")
    end

    # An OPC package: the object graph behind a .pptx file.
    class OpcPackage
      include Relatable

      # Open a package from a path, an IO stream, or a directory holding an
      # unzipped package.
      def self.open(pkg_file) = new(pkg_file).load

      def initialize(pkg_file = nil)
        @pkg_file = pkg_file
      end

      def rels = @rels ||= Relationships.new(PackURI::PACKAGE.base_uri)

      def drop_rel(r_id) = rels.delete(r_id)

      # The presentation part, which is the package's main document part.
      def main_document_part = part_related_by(RELATIONSHIP_TYPE::OFFICE_DOCUMENT)

      # Every part in the package, each yielded once.
      def parts
        return enum_for(:parts) unless block_given?

        seen = {}.compare_by_identity
        each_rel do |rel|
          next if rel.external?

          part = rel.target_part
          next if seen.key?(part)

          seen[part] = true
          yield part
        end
      end

      # Every relationship in the package, each yielded once, depth first.
      def each_rel(&block)
        return enum_for(:each_rel) unless block

        visited = {}.compare_by_identity
        walk = lambda do |relationships|
          relationships.each_value do |rel|
            block.call(rel)
            next if rel.external?

            part = rel.target_part
            next if visited.key?(part)

            visited[part] = true
            walk.call(part.rels)
          end
        end
        walk.call(rels)
      end

      # The next unused partname matching +template+, which contains a single
      # "%d", e.g. "/ppt/slides/slide%d.xml".
      def next_partname(template)
        prefix = template.split("%d").first
        taken = parts.each_with_object({}) do |part, names|
          name = part.partname.to_s
          names[name] = true if name.start_with?(prefix)
        end

        # Count down from the optimistic answer: in the common case every
        # number from 1 is in use and the first candidate is the right one.
        (taken.size + 1).downto(1) do |n|
          candidate = format(template, n)
          return PackURI.new(candidate) unless taken.key?(candidate)
        end
        raise Error, "ran out of candidate partnames for #{template}"
      end

      def save(pkg_file)
        PackageWriter.write(pkg_file, rels, parts.to_a)
        self
      end

      def load
        xml_rels, parts_by_name = PackageLoader.load(@pkg_file, self)
        rels.load_from_xml(PackURI::PACKAGE.base_uri, xml_rels, parts_by_name)
        self
      end
    end

    # Loads a package from its serialized form.
    class PackageLoader
      # @return [Array(Oxml::CT_Relationships, Hash{String=>Part})]
      def self.load(pkg_file, package) = new(pkg_file, package).load

      def initialize(pkg_file, package)
        @pkg_file = pkg_file
        @package = package
      end

      def load
        parts = build_parts
        parts.each { |partname, part| part.load_rels_from_xml(xml_rels[partname], parts) }
        [xml_rels[PackURI::PACKAGE.to_s], parts]
      end

      private

      def reader = @reader ||= PackageReader.new(@pkg_file)

      def content_types
        @content_types ||= ContentTypeMap.from_xml(reader[PackURI::CONTENT_TYPES])
      end

      # {partname => Part}, every part constructed but with relationships not
      # yet resolved -- that needs the whole collection to exist first.
      def build_parts
        xml_rels.keys.each_with_object({}) do |partname, parts|
          next if partname == PackURI::PACKAGE.to_s

          uri = PackURI.new(partname)
          # Some packages in the wild reference parts that are not present.
          # Skip them rather than refusing to open the file at all.
          next unless reader.key?(uri)

          parts[partname] = PartFactory.build(uri, content_types[uri], @package, reader[uri])
        end
      end

      # {partname => CT_Relationships} for the package and every part, found by
      # walking the relationship graph depth first from the package itself.
      def xml_rels
        @xml_rels ||= begin
          collected = {}
          visit = lambda do |source_partname, rels_element|
            collected[source_partname.to_s] = rels_element
            base_uri = source_partname.base_uri

            rels_element.relationship_list.each do |rel|
              next if rel.external?

              target = PackURI.from_rel_ref(base_uri, rel.target_ref)
              next if collected.key?(target.to_s)

              visit.call(target, xml_rels_for(target))
            end
          end
          visit.call(PackURI::PACKAGE, xml_rels_for(PackURI::PACKAGE))
          collected
        end
      end

      # A part with no relationships gets an empty <Relationships> element
      # rather than nil, so callers never branch on absence.
      def xml_rels_for(partname)
        rels_xml = reader.rels_xml_for(partname)
        rels_xml.nil? ? Oxml::CT_Relationships.new_element : Pptx::Oxml::Element.parse(rels_xml)
      end
    end

    # A part in the package. Subclassed to give particular content types their
    # own behaviour; used directly for parts with none.
    class Part
      include Relatable

      attr_reader :partname, :content_type, :package

      def self.load(partname, content_type, package, blob) = new(partname, content_type, package, blob)

      def initialize(partname, content_type, package, blob = nil)
        @partname = partname
        @content_type = content_type
        @package = package
        @blob = blob
      end

      def partname=(partname)
        unless partname.is_a?(PackURI)
          raise TypeError, "partname must be a PackURI, got #{partname.class}"
        end

        @partname = partname
      end

      # @return [String] the bytes this part serializes to
      def blob = @blob || ""

      attr_writer :blob

      def rels = @rels ||= Relationships.new(@partname.base_uri)

      def load_rels_from_xml(xml_rels, parts)
        rels.load_from_xml(@partname.base_uri, xml_rels, parts)
      end

      def drop_rel(r_id) = rels.delete(r_id)

      def inspect = "#<#{self.class.name} #{@partname}>"
    end

    # A part whose payload is XML, which is most of them.
    class XmlPart < Part
      attr_reader :element

      def self.load(partname, content_type, package, blob)
        new(partname, content_type, package, Pptx::Oxml::Element.parse(blob))
      end

      def initialize(partname, content_type, package, element)
        super(partname, content_type, package)
        @element = element
      end

      def blob = Oxml.serialize_part_xml(@element)

      # This part. Objects further down the tree delegate up to find the part
      # they live in; the chain ends here.
      def part = self

      # Drop a relationship unless the XML still refers to it more than once.
      #
      # A reference count of 0 means an implicit relationship, which is safe to
      # drop. Only XML parts can do this, since only they can be searched.
      def drop_rel(r_id)
        rels.delete(r_id) if rel_ref_count(r_id) < 2
      end

      private

      def rel_ref_count(r_id)
        @element.xpath("//@r:id").count { |attr| attr.value == r_id }
      end
    end

    # Builds the {Part} subclass registered for a content type.
    module PartFactory
      @part_type_for = {}

      class << self
        # Register +part_class+ as the class for +content_type+.
        def register(content_type, part_class)
          @part_type_for[content_type] = part_class
        end

        def part_class_for(content_type) = @part_type_for.fetch(content_type, Part)

        def registered = @part_type_for.dup

        # Test seam: restore a previous registration table.
        def reset!(to) = (@part_type_for = to)

        def build(partname, content_type, package, blob)
          part_class_for(content_type).load(partname, content_type, package, blob)
        end
      end
    end

    # Looks up the content type for a partname, per `[Content_Types].xml`.
    class ContentTypeMap
      def self.from_xml(content_types_xml)
        types = Pptx::Oxml::Element.parse(content_types_xml)
        overrides = types.override_list.to_h { |o| [o.partName.downcase, o.contentType] }
        defaults = types.default_list.to_h { |d| [d.extension.downcase, d.contentType] }
        new(overrides, defaults)
      end

      def initialize(overrides, defaults)
        @overrides = overrides
        @defaults = defaults
      end

      # Matching is case-insensitive: the spec says partnames and extensions
      # compare without regard to case.
      def [](partname)
        unless partname.is_a?(PackURI)
          raise TypeError, "content-type key must be a PackURI, got #{partname.class}"
        end

        @overrides[partname.to_s.downcase] ||
          @defaults[partname.ext.downcase] ||
          raise(KeyError, "no content-type for partname #{partname} in [Content_Types].xml")
      end
    end

    # The relationships belonging to one part, or to the package.
    #
    # Keyed by rId, and Enumerable over the relationships themselves.
    class Relationships
      include Enumerable

      def initialize(base_uri)
        @base_uri = base_uri
        @rels = {}
      end

      def each(&) = @rels.each_value(&)

      def each_value(&) = @rels.each_value(&)

      def [](r_id) = @rels[r_id]

      def fetch(r_id)
        @rels.fetch(r_id) { raise NotFoundError, "no relationship with key #{r_id.inspect}" }
      end

      def key?(r_id) = @rels.key?(r_id)
      def keys = @rels.keys
      def size = @rels.size
      def empty? = @rels.empty?
      def delete(r_id) = @rels.delete(r_id)

      # The rId of a relationship of +reltype+ to +target_part+, adding one if
      # no match exists.
      def get_or_add(reltype, target_part)
        matching(reltype, target_part, external: false) || add(reltype, target_part)
      end

      def get_or_add_external(reltype, target_ref)
        matching(reltype, target_ref, external: true) ||
          add(reltype, target_ref, external: true)
      end

      def part_with_reltype(reltype)
        matches = select { |rel| rel.reltype == reltype }
        raise NotFoundError, "no relationship of type #{reltype.inspect}" if matches.empty?

        if matches.size > 1
          raise Error, "multiple relationships of type #{reltype.inspect} in collection"
        end

        matches.first.target_part
      end

      # Replace the contents of this collection with the relationships in
      # +xml_rels+.
      def load_from_xml(base_uri, xml_rels, parts)
        @base_uri = base_uri
        @rels.clear
        xml_rels.relationship_list.each do |rel_elm|
          # A plugin sometimes "removes" a relationship by voiding its target
          # rather than deleting the element, leaving a link to something like
          # "/ppt/slides/NULL". Skip anything pointing outside the package.
          unless rel_elm.external?
            partname = PackURI.from_rel_ref(base_uri, rel_elm.target_ref)
            next unless parts.key?(partname.to_s)
          end

          rel = Relationship.from_xml(base_uri, rel_elm, parts)
          @rels[rel.r_id] = rel
        end
        self
      end

      # @return [String] serialized `.rels` bytes for this collection
      def xml
        rels_elm = Oxml::CT_Relationships.new_element
        in_numerical_order.each do |rel|
          rels_elm.add_rel(rel.r_id, rel.reltype, rel.target_ref, is_external: rel.external?)
        end
        rels_elm.xml_file_bytes
      end

      private

      # Emit <Relationship> elements in numerical rId order, so a saved package
      # is reproducible and diffable.
      def in_numerical_order
        @rels.keys.sort_by { |r_id| [numeric_part(r_id), r_id] }.map { |r_id| @rels[r_id] }
      end

      def numeric_part(r_id)
        match = /\ArId(\d+)\z/.match(r_id)
        match ? match[1].to_i : 0
      end

      def matching(reltype, target, external:)
        rel = find do |candidate|
          next false unless candidate.reltype == reltype && candidate.external? == external

          (external ? candidate.target_ref : candidate.target_part) == target
        end
        rel&.r_id
      end

      def add(reltype, target, external: false)
        r_id = next_r_id
        @rels[r_id] = Relationship.new(
          @base_uri, r_id, reltype, external ? RELATIONSHIP_TARGET_MODE::EXTERNAL
                                             : RELATIONSHIP_TARGET_MODE::INTERNAL, target
        )
        r_id
      end

      # The first unused rId from "rId1", filling any gap in the numbering.
      def next_r_id
        (@rels.size + 1).downto(1) do |n|
          candidate = "rId#{n}"
          return candidate unless @rels.key?(candidate)
        end
        raise Error, "unable to allocate an rId"
      end
    end

    # A link from a part (or the package) to another part or an external
    # resource.
    class Relationship
      attr_reader :r_id, :reltype

      def self.from_xml(base_uri, rel_elm, parts)
        target =
          if rel_elm.external?
            rel_elm.target_ref
          else
            parts.fetch(PackURI.from_rel_ref(base_uri, rel_elm.target_ref).to_s)
          end
        new(base_uri, rel_elm.rId, rel_elm.reltype, rel_elm.targetMode, target)
      end

      def initialize(base_uri, r_id, reltype, target_mode, target)
        @base_uri = base_uri
        @r_id = r_id
        @reltype = reltype
        @target_mode = target_mode
        @target = target
      end

      # True when the target is outside the package, such as a URL.
      def external? = @target_mode == RELATIONSHIP_TARGET_MODE::EXTERNAL

      def target_part
        raise Error, "#target_part is undefined for an external relationship" if external?

        @target
      end

      def target_partname
        raise Error, "#target_partname is undefined for an external relationship" if external?

        @target.partname
      end

      # The reference as written into the `.rels` item: a relative partname
      # internally, or the URL for an external relationship.
      def target_ref = external? ? @target : target_partname.relative_ref(@base_uri)

      def inspect = "#<Pptx::Opc::Relationship #{@r_id} #{@reltype}>"
    end
  end
end
