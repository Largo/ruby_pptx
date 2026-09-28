# frozen_string_literal: true

require "nokogiri"
require "ruby_pptx/errors"
require "ruby_pptx/oxml/ns"
require "ruby_pptx/oxml/simple_types"

module Pptx
  module Oxml
    # Maps a Clark-notation tag name to the {Element} subclass that models it.
    #
    # python-pptx gets this for free from lxml, whose parser instantiates a
    # custom class per tag. Nokogiri has no such hook, so element classes
    # register themselves here and {Element.wrap} does the dispatch.
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
        def class_for(clark_name) = @classes.fetch(clark_name, Element)

        def registered = @classes.dup

        # Test seam: forget registrations made by a spec.
        def reset!(to) = (@classes = to)
      end
    end

    # Wraps a `Nokogiri::XML::Node` and gives it the schema-aware behaviour of
    # one OOXML element type.
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

      def self.nsptag = @nsptags&.first

      class << self
        attr_reader :nsptags
      end

      # Wrap +node+ in the class registered for its tag.
      #
      # @param node [Nokogiri::XML::Node]
      # @return [Element, nil] nil when +node+ is nil
      def self.wrap(node)
        return nil if node.nil?
        return node if node.is_a?(Element)

        cache = cache_for(node.document)
        cache[node] ||= Registry.class_for(Ns.clark_name_of(node)).new(node)
      end

      # Parse +xml+ into a standalone element tree.
      def self.parse(xml)
        doc = Nokogiri::XML(xml) { |config| config.noblanks.strict }
        raise InvalidXmlError, doc.errors.first.to_s if doc.errors.any?

        wrap(doc.root)
      end

      # Per-document wrapper cache, keyed on node identity. Lives on the
      # document so it dies with it.
      def self.cache_for(document)
        document.instance_variable_get(:@pptx_wrappers) ||
          document.instance_variable_set(:@pptx_wrappers, {}.compare_by_identity)
      end
      private_class_method :cache_for

      # @return [Nokogiri::XML::Node] the wrapped node
      attr_reader :node

      def initialize(node)
        @node = node
      end

      # -- tree navigation -------------------------------------------------

      # @return [String] this element's prefixed tag, e.g. "p:sp"
      def nsptag = Ns.prefixed_tag(Ns.clark_name_of(@node))

      def document = @node.document

      def parent = Element.wrap(@node.parent.is_a?(Nokogiri::XML::Element) ? @node.parent : nil)

      # @return [Array<Element>] every child element, in document order
      def element_children = @node.element_children.map { |child| Element.wrap(child) }

      # First child element with +nsptag+, or nil.
      def find(nsptag) = Element.wrap(raw_find(nsptag))

      # Every child element with +nsptag+, in document order.
      def find_all(nsptag)
        pfx, local = Ns.split_tag(nsptag)
        @node.element_children
             .select { |c| c.name == local && c.namespace&.href == Ns.nsuri(pfx) }
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
        @node.xpath(expression, Ns::NSMAP).map do |result|
          result.is_a?(Nokogiri::XML::Element) ? Element.wrap(result) : result
        end
      end

      # -- mutation --------------------------------------------------------

      # Insert +element+ before the first of +successor_tags+ present, or
      # append it when none is. This is what keeps children in schema order.
      def insert_element_before(element, *successor_tags)
        successor = first_child_found_in(*successor_tags)
        if successor
          successor.node.add_previous_sibling(adopt(element))
        else
          @node.add_child(adopt(element))
        end
        element
      end

      def append(element)
        @node.add_child(adopt(element))
        element
      end

      # Remove every child element whose tag is in +nsptags+.
      def remove_all(*nsptags)
        nsptags.each { |nsptag| find_all(nsptag).each { |child| child.node.unlink } }
        nil
      end

      def remove(element)
        element.node.unlink
        nil
      end

      # Create a new, empty element with +nsptag+, in *this* element's
      # document.
      #
      # Creating it anywhere else would be a bug: Nokogiri re-creates a node
      # adopted from another document, so the object you built is not the
      # object that ends up in the tree, and every wrapper held on it goes
      # stale. Staying in one document keeps node identity intact.
      def build(nsptag)
        pfx, local = Ns.split_tag(nsptag)
        uri = Ns.nsuri(pfx)
        child = Nokogiri::XML::Node.new(local, @node.document)
        # Resolve against this element's scope, not the detached child's, so
        # the new element joins an existing declaration instead of introducing
        # a second one. Where the parent holds the namespace as its *default*
        # -- as .rels and [Content_Types].xml do -- the child is created
        # unprefixed to match, which matters because C14N preserves prefixes.
        child.namespace = inherited_namespace(pfx, uri) ||
                          child.add_namespace_definition(pfx, uri)
        Element.wrap(child)
      end

      # Parse an XML literal into this element's document.
      def build_from_xml(xml)
        # noblanks matches Element.parse, so an indented XML literal does not
        # smuggle whitespace text nodes into the tree.
        fragment = @node.parse(xml, &:noblanks)
        raise InvalidXmlError, "fragment produced no element: #{xml.inspect}" if fragment.first.nil?

        Element.wrap(fragment.first)
      end

      # -- text content ----------------------------------------------------

      # @return [String] the element's text content
      def text = @node.text

      def text=(value)
        @node.content = value.to_s
      end

      # Declare +prefix+ on this element even if nothing uses it yet.
      #
      # Needed where a descendant carries a namespaced attribute but the
      # declaration belongs on the root, which is where PowerPoint puts it.
      def declare_namespace(prefix)
        namespace_for(@node, prefix)
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
        return @node[attr_name.to_s] unless pfx

        @node.attribute_with_ns(local, Ns.nsuri(pfx))&.value
      end

      def set(attr_name, value)
        pfx, local = prefixed_attr_parts(attr_name)
        return @node[attr_name.to_s] = value.to_s unless pfx

        uri = Ns.nsuri(pfx)
        if (existing = @node.attribute_with_ns(local, uri))
          existing.value = value.to_s
        else
          namespace_for(@node, pfx)
          @node["#{pfx}:#{local}"] = value.to_s
        end
        value
      end

      def delete_attribute(attr_name)
        pfx, local = prefixed_attr_parts(attr_name)
        return @node.remove_attribute(attr_name.to_s) unless pfx

        @node.attribute_with_ns(local, Ns.nsuri(pfx))&.unlink
        nil
      end

      def attribute?(attr_name) = !get(attr_name).nil?

      # -- serialization and equality --------------------------------------

      # Pretty-printed XML for this element, without a declaration. For
      # debugging and specs.
      def xml = @node.to_xml(indent: 2)

      def to_s = xml

      def ==(other) = other.is_a?(Element) && other.node.equal?(@node)
      alias eql? ==

      def hash = @node.object_id.hash

      def inspect = "#<#{self.class.name} <#{nsptag}> children=#{@node.element_children.size}>"

      private

      def raw_find(nsptag)
        pfx, local = Ns.split_tag(nsptag)
        uri = Ns.nsuri(pfx)
        @node.element_children.find { |c| c.name == local && c.namespace&.href == uri }
      end

      # Guard against the cross-document trap described on {#build}.
      def adopt(element)
        node = element.is_a?(Element) ? element.node : element
        return node if node.document.equal?(@node.document)

        raise ArgumentError,
              "cannot insert an element created in a different document; " \
              "build it with #build on an element of this tree"
      end

      def prefixed_attr_parts(attr_name)
        return nil unless attr_name.to_s.include?(":")

        Ns.split_tag(attr_name.to_s)
      end

      # An in-scope declaration for +uri+, preferring one bound to +prefix+ but
      # accepting any -- including a default (unprefixed) declaration.
      def inherited_namespace(prefix, uri)
        scopes = @node.namespace_scopes
        scopes.find { |ns| ns.href == uri && ns.prefix == prefix } ||
          scopes.find { |ns| ns.href == uri }
      end

      # Reuse an in-scope namespace declaration when there is one, so the
      # output does not sprout a redundant xmlns on every new element.
      def namespace_for(node, prefix)
        uri = Ns.nsuri(prefix)
        existing = node.namespace_scopes.find { |ns| ns.href == uri && ns.prefix == prefix }
        existing || node.add_namespace_definition(prefix, uri)
      end
    end
  end
end
