# frozen_string_literal: true

require "zip"
require "stringio"
require "ruby_pptx/errors"
require "ruby_pptx/opc/constants"
require "ruby_pptx/opc/oxml"
require "ruby_pptx/opc/pack_uri"
require "ruby_pptx/opc/spec"

module Pptx
  module Opc
    # Reads package items out of a serialized package.
    #
    # The package may be a .pptx file, an IO stream containing one, or a
    # directory holding an unzipped package.
    class PackageReader
      def initialize(pkg_file)
        @reader = PhysicalReader.for(pkg_file)
      end

      # @return [Boolean] whether the package holds an item at +pack_uri+
      def key?(pack_uri) = @reader.key?(pack_uri)

      # @return [String] binary contents of the item at +pack_uri+
      # @raise [KeyError] when there is no such item
      def [](pack_uri) = @reader[pack_uri]

      # The XML of the `.rels` item belonging to +partname+, or nil when that
      # part has no relationships.
      def rels_xml_for(partname)
        uri = partname.rels_uri
        key?(uri) ? self[uri] : nil
      end
    end

    # Writes a zip-format package.
    class PackageWriter
      # Write a .pptx containing +parts+, their rels items, the package rels
      # item and a content-types stream derived from the parts.
      def self.write(pkg_file, pkg_rels, parts)
        new(pkg_file, pkg_rels, parts).write
      end

      def initialize(pkg_file, pkg_rels, parts)
        @pkg_file = pkg_file
        @pkg_rels = pkg_rels
        @parts = parts
      end

      def write
        PhysicalWriter.open(@pkg_file) do |writer|
          writer.write(PackURI::CONTENT_TYPES, ContentTypesItem.xml_for(@parts))
          writer.write(PackURI::PACKAGE.rels_uri, @pkg_rels.xml)
          @parts.each do |part|
            writer.write(part.partname, part.blob)
            writer.write(part.partname.rels_uri, part.rels.xml) unless part.rels.empty?
          end
        end
      end
    end

    # Reads items from a physical package. {PhysicalReader.for} picks the
    # implementation matching what it is handed.
    class PhysicalReader
      def self.for(pkg_file)
        return ZipReader.new(pkg_file) unless pkg_file.is_a?(String)
        return DirectoryReader.new(pkg_file) if File.directory?(pkg_file)
        return ZipReader.new(pkg_file) if File.file?(pkg_file) && zip?(pkg_file)

        raise PackageNotFoundError, "package not found at #{pkg_file.inspect}"
      end

      def self.zip?(path)
        File.binread(path, 4) == "PK\x03\x04".b
      rescue SystemCallError
        false
      end
      private_class_method :zip?
    end

    # Reads a package that has been unzipped into a directory.
    class DirectoryReader < PhysicalReader
      def initialize(path)
        super()
        @path = File.expand_path(path)
      end

      def key?(pack_uri) = File.file?(path_for(pack_uri))

      def [](pack_uri)
        File.binread(path_for(pack_uri))
      rescue SystemCallError
        raise KeyError, "no member #{pack_uri} in package"
      end

      private

      # Part names are package-internal and must not be able to escape the
      # directory, so a name that resolves outside it is refused.
      def path_for(pack_uri)
        candidate = File.expand_path(File.join(@path, pack_uri.member_name))
        unless candidate == @path || candidate.start_with?("#{@path}/")
          raise ArgumentError, "part name escapes the package directory: #{pack_uri}"
        end

        candidate
      end
    end

    # Reads a zip-format package, given a path or an IO stream.
    class ZipReader < PhysicalReader
      def initialize(pkg_file)
        super()
        @pkg_file = pkg_file
      end

      def key?(pack_uri) = blobs.key?(pack_uri.to_s)

      def [](pack_uri)
        blobs.fetch(pack_uri.to_s) { raise KeyError, "no member #{pack_uri} in package" }
      end

      private

      # Read every member up front: a package is traversed by relationship, not
      # sequentially, and the whole thing is needed anyway.
      def blobs
        @blobs ||= each_entry.to_h { |name, data| ["/#{name}", data] }
      end

      def each_entry(&block)
        return enum_for(:each_entry) unless block

        open_zip do |zip|
          zip.each do |entry|
            block.call(entry.name, entry.get_input_stream.read) if entry.file?
          end
        end
      end

      def open_zip(&)
        if @pkg_file.is_a?(String)
          Zip::File.open(@pkg_file, &)
        else
          @pkg_file.rewind if @pkg_file.respond_to?(:rewind)
          Zip::File.open_buffer(@pkg_file, &)
        end
      rescue Zip::Error => e
        raise PackageNotFoundError, "not a readable OPC package: #{e.message}"
      end
    end

    # Writes a zip-format package to a path or an IO stream.
    class PhysicalWriter
      def self.open(pkg_file, &)
        ZipWriter.open(pkg_file, &)
      end
    end

    class ZipWriter < PhysicalWriter
      def self.open(pkg_file)
        if pkg_file.is_a?(String)
          Zip::OutputStream.open(pkg_file) { |zos| yield new(zos) }
        else
          buffer = Zip::OutputStream.write_buffer(StringIO.new) { |zos| yield new(zos) }
          buffer.rewind
          IO.copy_stream(buffer, pkg_file)
        end
      end

      def initialize(zip_output_stream)
        super()
        @zos = zip_output_stream
      end

      def write(pack_uri, blob)
        @zos.put_next_entry(pack_uri.member_name)
        @zos.write(blob)
      end
    end

    # Composes `[Content_Types].xml` for a collection of parts.
    class ContentTypesItem
      # @return [String] serialized content-types XML covering every part
      def self.xml_for(parts)
        Oxml.serialize_part_xml(new(parts).element)
      end

      def initialize(parts)
        @parts = parts
      end

      # Defaults are sorted by extension and overrides by partname. The order
      # carries no meaning in the spec, but a deterministic one makes packages
      # diffable and testable.
      def element
        defaults, overrides = defaults_and_overrides
        types = Oxml::CT_Types.new_element
        defaults.sort.each { |ext, content_type| types.add_default_for(ext, content_type) }
        overrides.sort.each { |partname, content_type| types.add_override_for(partname, content_type) }
        types
      end

      private

      # A part gets a <Default> when its (extension, content-type) pair is one
      # the spec allows to be defaulted, and an <Override> otherwise.
      def defaults_and_overrides
        defaults = {
          "rels" => CONTENT_TYPE::OPC_RELATIONSHIPS,
          "xml" => CONTENT_TYPE::XML
        }
        overrides = {}

        @parts.each do |part|
          ext = part.partname.ext
          if DEFAULT_CONTENT_TYPES.include?([ext.downcase, part.content_type])
            defaults[ext.downcase] = part.content_type
          else
            overrides[part.partname.to_s] = part.content_type
          end
        end

        [defaults, overrides]
      end
    end
  end
end
