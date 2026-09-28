# frozen_string_literal: true

require "ruby_pptx/errors"
require "ruby_pptx/oxml/backend"
require "ruby_pptx/oxml/ns"
require "ruby_pptx/oxml/simple_types"

module Pptx
  module Oxml
    # The XML library in use, chosen once at load; see {Backend}.
    BACKEND = Backend.select

    # Maps a Clark-notation tag name to the {Element} subclass that models it.
    #
    # python-pptx gets this for free from lxml, whose parser instantiates a
    # custom class per tag. Neither Nokogiri nor REXML has such a hook, so
    # element classes register themselves here and {Element.wrap} does the
    # dispatch.
    module Registry
      @classes = {}

      class << self
        # A tag belongs to exactly one class. Registering a second one used to
        # replace the first silently, so whichever file loaded last won --
        # `c:tx` and `a:ext` each had two classes that way.
        def register(clark_name, element_class)
          existing = @classes[clark_name]
          if existing && existing != element_class
            raise ArgumentError, "#{clark_name} is already registered to #{existing}"
          end

          @classes[clark_name] = element_class
        end

        # @return [Class] the registered class, or {Element} for an unmodelled tag
        def class_for(clark_name)
          @classes.fetch(clark_name, Element)
        end

        def registered
          @classes.dup
        end

        # Test seam: forget registrations made by a spec.
        def reset!(to)
          (@classes = to)
        end
      end
    end

    # Wraps an XML node -- Nokogiri's or REXML's, see {Backend} -- and gives
    # it the schema-aware behaviour of one OOXML element type.
    #
    # Subclasses declare their content model with class macros, which generate
    # the accessors that the rest of the library uses:
    #
    #   class CT_Shape < Element
    #     tag "p:sp"
    #     one_and_only_one "p:nvSpPr"
    #     zero_or_one      "p:spPr", successors: %w[p:style p:txBody]
    #     zero_or_more     "p:ext"
    #     optional_attr    "macro", type: SimpleTypes::XsdString
    #   end
    #
    # Local tag names are kept verbatim, camel case and all, so a method name
    # can be grepped straight from the ECMA-376 schema. The snake_case Ruby
    # idiom belongs to the public API layer built on top of this one.
    #
    # ## Identity
    #
    # Two wrappers around the same underlying node are `==` and hash alike, so
    # correctness never depends on the wrapper cache. The cache exists to avoid
    # re-allocating wrappers, not to make comparisons work.
    class Element
      # Declares the tag or tags this class models and registers them for
      # dispatch. A few element types appear under more than one tag -- a
      # transform is `a:xfrm` on a shape but `p:xfrm` on a graphic frame --
      # and are the same class in both places.
      def self.tag(*prefixed_tags)
        @nsptags = prefixed_tags
        prefixed_tags.each { |t| Registry.register(Ns.qn(t), self) }
      end

      def self.nsptag
        @nsptags&.first
      end

      class << self
        attr_reader :nsptags
      end

      # Wrap +node+ in the class registered for its tag.
      #
      # @param node [Object] a node of the active {Backend}
      # @return [Element, nil] nil when +node+ is nil
      def self.wrap(node)
        return nil if node.nil?
        return node if node.is_a?(Element)

        cache = BACKEND.wrapper_cache(node)
        cache[node] ||= Registry.class_for(Ns.clark_name_of(node)).new(node)
      end

      # Parse +xml+ into a standalone element tree.
      def self.parse(xml)
        wrap(BACKEND.parse(xml))
      end

      # @return [Object] the wrapped node, a Nokogiri or REXML element
      #   depending on the active {Backend}. An escape hatch: code that uses
      #   it ties itself to one backend.
      attr_reader :node

      def initialize(node)
        @node = node
      end

      # -- tree navigation -------------------------------------------------

      # @return [String] this element's prefixed tag, e.g. "p:sp"
      def nsptag
        Ns.prefixed_tag(Ns.clark_name_of(@node))
      end

      # An opaque token for the document this element belongs to; equal
      # (+equal?+) for two elements exactly when they share a document.
      def document
        BACKEND.document(@node)
      end

      def parent
        Element.wrap(BACKEND.parent_element(@node))
      end

      # @return [Array<Element>] every child element, in document order
      def element_children
        BACKEND.element_children(@node).map { |child| Element.wrap(child) }
      end

      # @return [String] this element's tag in Clark notation, "{uri}local"
      def clark_name
        Ns.clark_name_of(@node)
      end

      # First child element with +nsptag+, or nil.
      def find(nsptag)
        Element.wrap(raw_find(nsptag))
      end

      # Every child element with +nsptag+, in document order.
      def find_all(nsptag)
        pfx, local = Ns.split_tag(nsptag)
        uri = Ns.nsuri(pfx)
        BACKEND.element_children(@node)
               .select { |c| BACKEND.local_name(c) == local && BACKEND.namespace_uri(c) == uri }
               .map { |c| Element.wrap(c) }
      end

      # The first child matching any of +nsptags+, in the order given -- not in
      # document order. Used to find an element's successor when inserting.
      def first_child_found_in(*nsptags)
        nsptags.each do |nsptag|
          child = raw_find(nsptag)
          return Element.wrap(child) if child
        end
        nil
      end

      # Run an XPath expression with the full OOXML namespace map bound.
      def xpath(expression)
        BACKEND.xpath(@node, expression, Ns::NSMAP).map do |result|
          BACKEND.element?(result) ? Element.wrap(result) : result
        end
      end

      # -- mutation --------------------------------------------------------

      # Insert +element+ before the first of +successor_tags+ present, or
      # append it when none is. This is what keeps children in schema order.
      def insert_element_before(element, *successor_tags)
        successor = first_child_found_in(*successor_tags)
        if successor
          successor.add_previous_sibling(element)
        else
          append(element)
        end
        element
      end

      def append(element)
        BACKEND.append_child(@node, adopt(element))
        element
      end

      # Make +element+ this element's first child.
      def prepend(element)
        BACKEND.prepend_child(@node, adopt(element))
        element
      end

      # Put +element+ immediately before this one.
      def add_previous_sibling(element)
        BACKEND.insert_before(@node, adopt(element))
        element
      end

      # Put +element+ immediately after this one.
      def add_next_sibling(element)
        BACKEND.insert_after(@node, adopt(element))
        element
      end

      # Remove every child element whose tag is in +nsptags+.
      def remove_all(*nsptags)
        nsptags.each { |nsptag| find_all(nsptag).each { |child| BACKEND.unlink(child.node) } }
        nil
      end

      def remove(element)
        BACKEND.unlink(element.node)
        nil
      end

      # Create a new, empty element with +nsptag+, in *this* element's
      # document, reusing a namespace declaration already in scope here.
      #
      # Build new elements this way rather than parsing them separately:
      # Nokogiri re-creates a node adopted from another document, so the
      # object built is not the object that ends up in the tree, and every
      # wrapper held on it goes stale.
      def build(nsptag)
        pfx, local = Ns.split_tag(nsptag)
        Element.wrap(BACKEND.build(@node, pfx, local, Ns.nsuri(pfx)))
      end

      # Parse an XML literal into this element's document, its prefixes
      # resolved against the namespaces in scope here.
      def build_from_xml(xml)
        fragment = BACKEND.parse_fragment(@node, xml)
        raise InvalidXmlError, "fragment produced no element: #{xml.inspect}" if fragment.nil?

        Element.wrap(fragment)
      end

      # A copy of +element+, from any document, built into this one.
      def import(element)
        build_from_xml(element.to_xml)
      end

      # -- text content ----------------------------------------------------

      # @return [String] the element's text content
      def text
        BACKEND.text(@node)
      end

      def text=(value)
        BACKEND.set_text(@node, value.to_s)
      end

      # Declare +prefix+ on this element even if nothing uses it yet.
      #
      # Needed where a descendant carries a namespaced attribute but the
      # declaration belongs on the root, which is where PowerPoint puts it.
      def declare_namespace(prefix)
        BACKEND.declare_namespace(@node, prefix, Ns.nsuri(prefix))
        self
      end

      # -- attributes ------------------------------------------------------

      # Attribute names may be prefixed ("r:embed") or not ("cstate").
      #
      # A prefixed name is resolved by namespace URI rather than by the literal
      # prefix, because a document is free to bind that namespace to any prefix
      # it likes. This matches the Clark-name semantics python-pptx inherits
      # from lxml.
      def get(attr_name)
        pfx, local = prefixed_attr_parts(attr_name)
        return BACKEND.get_attribute(@node, attr_name.to_s) unless pfx

        BACKEND.get_attribute_ns(@node, local, Ns.nsuri(pfx))
      end

      def set(attr_name, value)
        pfx, local = prefixed_attr_parts(attr_name)
        if pfx
          BACKEND.set_attribute_ns(@node, pfx, local, Ns.nsuri(pfx), value.to_s)
        else
          BACKEND.set_attribute(@node, attr_name.to_s, value.to_s)
        end
        value
      end

      def delete_attribute(attr_name)
        pfx, local = prefixed_attr_parts(attr_name)
        if pfx
          BACKEND.remove_attribute_ns(@node, local, Ns.nsuri(pfx))
        else
          BACKEND.remove_attribute(@node, attr_name.to_s)
        end
        nil
      end

      def attribute?(attr_name)
        !get(attr_name).nil?
      end

      # -- serialization and equality --------------------------------------

      # Pretty-printed XML for this element, without a declaration. For
      # debugging and specs.
      def xml
        BACKEND.pretty_xml(@node)
      end

      # Compact XML for this element, without a declaration.
      def to_xml
        BACKEND.to_xml(@node)
      end

      def to_s
        xml
      end

      def ==(other)
        other.is_a?(Element) && other.node.equal?(@node)
      end
      alias eql? ==

      def hash
        @node.object_id.hash
      end

      def inspect
        "#<#{self.class.name} <#{nsptag}> children=#{BACKEND.element_children(@node).size}>"
      end

      private

      def raw_find(nsptag)
        pfx, local = Ns.split_tag(nsptag)
        uri = Ns.nsuri(pfx)
        BACKEND.element_children(@node).find do |c|
          BACKEND.local_name(c) == local && BACKEND.namespace_uri(c) == uri
        end
      end

      # Guard against the cross-document trap described on {#build}.
      def adopt(element)
        node = element.is_a?(Element) ? element.node : element
        return node if BACKEND.document(node).equal?(document)

        raise ArgumentError,
              "cannot insert an element created in a different document; " \
              "build it with #build on an element of this tree"
      end

      def prefixed_attr_parts(attr_name)
        return nil unless attr_name.to_s.include?(":")

        Ns.split_tag(attr_name.to_s)
      end
    end
  end
end
