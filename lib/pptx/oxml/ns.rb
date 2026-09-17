# frozen_string_literal: true

module Pptx
  module Oxml
    # The XML namespaces used across PresentationML, DrawingML and the OPC
    # package format, plus helpers for moving between the three tag notations
    # we need:
    #
    # - prefixed, `"p:cSld"`               -- how the spec and this code talk
    # - Clark, `"{http://...}cSld"`        -- unambiguous, used as registry keys
    # - Nokogiri, name + namespace object  -- what the parser hands back
    module Ns
      NSMAP = {
        "a" => "http://schemas.openxmlformats.org/drawingml/2006/main",
        "c" => "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "cp" => "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "ct" => "http://schemas.openxmlformats.org/package/2006/content-types",
        "dc" => "http://purl.org/dc/elements/1.1/",
        "dcmitype" => "http://purl.org/dc/dcmitype/",
        "dcterms" => "http://purl.org/dc/terms/",
        "ep" => "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
        "i" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        "m" => "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "mo" => "http://schemas.microsoft.com/office/mac/office/2008/main",
        "mv" => "urn:schemas-microsoft-com:mac:vml",
        "o" => "urn:schemas-microsoft-com:office:office",
        "p" => "http://schemas.openxmlformats.org/presentationml/2006/main",
        # PowerPoint 2010 extensions, which is where slide sections live.
        "p14" => "http://schemas.microsoft.com/office/powerpoint/2010/main",
        "pd" => "http://schemas.openxmlformats.org/drawingml/2006/presentationDrawing",
        "pic" => "http://schemas.openxmlformats.org/drawingml/2006/picture",
        "pr" => "http://schemas.openxmlformats.org/package/2006/relationships",
        "r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "sl" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        "v" => "urn:schemas-microsoft-com:vml",
        "ve" => "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w10" => "urn:schemas-microsoft-com:office:word",
        "wne" => "http://schemas.microsoft.com/office/word/2006/wordml",
        "wp" => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "xsi" => "http://www.w3.org/2001/XMLSchema-instance"
      }.freeze

      # Namespace URI => preferred prefix. Built from NSMAP, so a URI that
      # appears twice resolves to the last prefix declared for it.
      PFXMAP = NSMAP.each_with_object({}) { |(pfx, uri), map| map[uri] = pfx }.freeze

      module_function

      # @param nspfx [String] e.g. "p"
      # @return [String] the namespace URI for +nspfx+
      def nsuri(nspfx)
        NSMAP.fetch(nspfx) { raise KeyError, "unknown namespace prefix #{nspfx.inspect}" }
      end

      # Clark-notation name for a prefixed tag.
      #
      #   qn("p:cSld") #=> "{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"
      def qn(prefixed_tag)
        pfx, local = split_tag(prefixed_tag)
        "{#{nsuri(pfx)}}#{local}"
      end

      # The inverse of {qn}.
      #
      #   prefixed_tag("{http://...presentationml/2006/main}cSld") #=> "p:cSld"
      def prefixed_tag(clark_name)
        unless clark_name.start_with?("{") && (close = clark_name.index("}"))
          raise ArgumentError, "not a Clark-notation name: #{clark_name.inspect}"
        end

        uri   = clark_name[1...close]
        local = clark_name[(close + 1)..]
        pfx = PFXMAP.fetch(uri) { raise KeyError, "unknown namespace URI #{uri.inspect}" }
        "#{pfx}:#{local}"
      end

      # The subset of NSMAP for +prefixes+, in the shape Nokogiri's #xpath wants.
      #
      #   namespaces("a", "p") #=> {"a" => "...", "p" => "..."}
      def namespaces(*prefixes)
        prefixes.to_h { |pfx| [pfx, nsuri(pfx)] }
      end

      # Namespace declarations as literal XML attribute text, for building
      # fragments by hand.
      #
      #   nsdecls("a") #=> 'xmlns:a="http://...drawingml/2006/main"'
      def nsdecls(*prefixes)
        prefixes.map { |pfx| %(xmlns:#{pfx}="#{nsuri(pfx)}") }.join(" ")
      end

      # The Clark name of a Nokogiri node, suitable as an element-registry key.
      #
      # @param node [Nokogiri::XML::Node]
      def clark_name_of(node)
        href = node.namespace&.href
        href ? "{#{href}}#{node.name}" : node.name
      end

      # ["p", "cSld"] for "p:cSld".
      def split_tag(prefixed_tag)
        pfx, local, extra = prefixed_tag.split(":", 3)
        if local.nil? || local.empty? || !extra.nil?
          raise ArgumentError, "expected a prefixed tag like 'p:cSld', got #{prefixed_tag.inspect}"
        end

        [pfx, local]
      end
    end
  end
end
