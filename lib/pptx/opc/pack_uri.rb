# frozen_string_literal: true

module Pptx
  module Opc
    # An absolute part name within an OPC package, e.g. "/ppt/slides/slide1.xml".
    #
    # Part names are always rooted at "/" and are *not* filesystem paths: they
    # are normalised purely lexically, with no reference to the working
    # directory or to `~`. The package itself has the pseudo part name "/".
    class PackURI
      include Comparable

      # Splits a filename stem into its name and optional trailing index, e.g.
      # "slide21" => ["slide", "21"].
      FILENAME_RE = /\A([a-zA-Z]+)([0-9]+)?/

      # Resolve +relative_ref+ (as found in a .rels part) against +base_uri+.
      #
      #   PackURI.from_rel_ref("/ppt", "slides/slide1.xml") #=> "/ppt/slides/slide1.xml"
      #   PackURI.from_rel_ref("/ppt/slides", "../media/image1.png") #=> "/ppt/media/image1.png"
      def self.from_rel_ref(base_uri, relative_ref)
        new(normalize(join(base_uri, relative_ref)))
      end

      # Lexically join path segments with "/", ignoring a leading "/" on later
      # segments the way posixpath.join does not -- we only ever join a base
      # with a genuinely relative reference.
      def self.join(*parts)
        parts.reject { |p| p.nil? || p.empty? }
             .each_with_index
             .map { |p, i| i.zero? ? p.chomp("/") : p }
             .join("/")
      end
      private_class_method :join

      # Collapse "." and ".." segments. Leading ".." on an absolute path are
      # dropped, matching posixpath.abspath.
      def self.normalize(path)
        absolute = path.start_with?("/")
        segments = path.split("/").each_with_object([]) do |seg, acc|
          case seg
          when "", "." then next
          when ".."
            if acc.empty? || acc.last == ".."
              acc.push(seg) unless absolute
            else
              acc.pop
            end
          else acc.push(seg)
          end
        end
        absolute ? "/#{segments.join("/")}" : segments.join("/")
      end
      private_class_method :normalize

      # @return [String] the full part name, e.g. "/ppt/slides/slide1.xml"
      attr_reader :to_s

      def initialize(pack_uri_str)
        pack_uri_str = pack_uri_str.to_s
        unless pack_uri_str.start_with?("/")
          raise ArgumentError, "PackURI must begin with a slash, got #{pack_uri_str.inspect}"
        end

        @to_s = pack_uri_str.freeze
        freeze
      end

      # The directory portion, e.g. "/ppt/slides" for "/ppt/slides/slide1.xml".
      # The package pseudo part name "/" has itself as its base URI.
      def base_uri
        idx = @to_s.rindex("/")
        idx.zero? ? "/" : @to_s[0...idx]
      end

      # The extension without its leading period, e.g. "xml". Empty when there
      # is none.
      def ext
        File.extname(filename).delete_prefix(".")
      end

      # The last segment, e.g. "slide1.xml". Empty for the package itself.
      def filename
        @to_s[(@to_s.rindex("/") + 1)..]
      end

      # The integer suffix of an "array" part name, e.g. 21 for
      # "/ppt/slides/slide21.xml"; nil for a singleton such as
      # "/ppt/presentation.xml".
      #
      # @return [Integer, nil]
      def idx
        stem = File.basename(filename, ".*")
        return nil if stem.empty?

        match = FILENAME_RE.match(stem)
        match && match[2] && Integer(match[2], 10)
      end

      # The part name without its leading slash -- the form used as the Zip
      # entry name.
      def member_name = @to_s[1..]

      # A reference to this part relative to +base_uri+, as written into a
      # .rels part.
      #
      #   PackURI.new("/ppt/slideLayouts/slideLayout1.xml").relative_ref("/ppt/slides")
      #   #=> "../slideLayouts/slideLayout1.xml"
      def relative_ref(base_uri)
        return member_name if base_uri == "/"

        from = base_uri.split("/").reject(&:empty?)
        to   = @to_s.split("/").reject(&:empty?)
        common = 0
        common += 1 while common < from.length && common < to.length && from[common] == to[common]
        (([".."] * (from.length - common)) + to[common..]).join("/")
      end

      # The part name of the .rels part describing this part's relationships.
      #
      #   PackURI.new("/ppt/presentation.xml").rels_uri #=> "/ppt/_rels/presentation.xml.rels"
      def rels_uri
        PackURI.new(self.class.send(:join, base_uri, "_rels", "#{filename}.rels"))
      end

      def <=>(other)
        other = PackURI.new(other) if other.is_a?(String)
        return nil unless other.is_a?(PackURI)

        @to_s <=> other.to_s
      end

      def ==(other) = (self <=> other)&.zero? || false
      alias eql? ==

      def hash = @to_s.hash

      def to_str = @to_s

      def inspect = "#<Pptx::Opc::PackURI #{@to_s.inspect}>"

      PACKAGE       = new("/")
      CONTENT_TYPES = new("/[Content_Types].xml")
    end
  end
end
