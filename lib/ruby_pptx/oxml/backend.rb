# frozen_string_literal: true

require "ruby_pptx/errors"

module Pptx
  module Oxml
    # The XML library underneath {Element}.
    #
    # Nokogiri is fast but native, so it cannot run everywhere -- ruby.wasm in
    # particular. REXML is pure Ruby and runs anywhere Ruby does, at a cost in
    # speed. {Element} talks to whichever is active through the functions
    # below and never touches a node directly; everything above {Element}
    # sees only elements.
    #
    # Nokogiri is used when it can be loaded, REXML otherwise. Set
    # `RUBY_PPTX_XML_BACKEND` to `nokogiri` or `rexml` to choose explicitly;
    # it is read once, when the library loads, because nodes from the two
    # cannot be mixed in one document.
    module Backend
      ENV_VAR = "RUBY_PPTX_XML_BACKEND"

      # @return [Module] {NokogiriBackend} or {RexmlBackend}
      def self.select(requested = ENV.fetch(ENV_VAR, nil))
        case requested&.downcase
        when nil, ""
          nokogiri_available? ? NokogiriBackend : RexmlBackend.tap(&:load)
        when "nokogiri" then NokogiriBackend.tap { require "nokogiri" }
        when "rexml" then RexmlBackend.tap(&:load)
        else raise ArgumentError, "#{ENV_VAR} must be nokogiri or rexml, not #{requested.inspect}"
        end
      end

      def self.nokogiri_available?
        require "nokogiri"
        true
      rescue LoadError
        false
      end
      private_class_method :nokogiri_available?

      # Escaping shared by both serializers where they must agree: the
      # characters libxml2 escapes, so REXML output matches Nokogiri's.
      ATTR_ESCAPES = { "&" => "&amp;", "<" => "&lt;", ">" => "&gt;", '"' => "&quot;",
                       "\n" => "&#10;", "\r" => "&#13;", "\t" => "&#9;" }.freeze
      TEXT_ESCAPES = { "&" => "&amp;", "<" => "&lt;", ">" => "&gt;", "\r" => "&#13;" }.freeze

      # -----------------------------------------------------------------------
      # Nokogiri
      # -----------------------------------------------------------------------
      module NokogiriBackend
        module_function

        def name
          :nokogiri
        end

        def parse(xml)
          doc = Nokogiri::XML(xml) { |config| config.noblanks.strict }
          raise InvalidXmlError, doc.errors.first.to_s if doc.errors.any?

          doc.root
        end

        # Parse +xml+ in the context of +node+, so prefixes resolve against the
        # namespaces in scope there. noblanks matches {parse}, so an indented
        # XML literal does not smuggle whitespace text nodes into the tree.
        def parse_fragment(node, xml)
          node.parse(xml, &:noblanks).first
        end

        # A new, detached element in +context+'s document.
        #
        # Creating it anywhere else would be a bug: Nokogiri re-creates a node
        # adopted from another document, so the object built is not the object
        # that ends up in the tree, and every wrapper held on it goes stale.
        #
        # Resolve against +context+'s scope, not the detached child's, so the
        # new element joins an existing declaration instead of introducing a
        # second one. Where the context holds the namespace as its *default*
        # -- as .rels and [Content_Types].xml do -- the child is created
        # unprefixed to match, which matters because C14N preserves prefixes.
        def build(context, prefix, local, uri)
          child = Nokogiri::XML::Node.new(local, context.document)
          scopes = context.namespace_scopes
          inherited = scopes.find { |ns| ns.href == uri && ns.prefix == prefix } ||
                      scopes.find { |ns| ns.href == uri }
          child.namespace = inherited || child.add_namespace_definition(prefix, uri)
          child
        end

        def document(node)
          node.document
        end

        # Per-document wrapper cache, keyed on node identity. Lives on the
        # document so it dies with it.
        def wrapper_cache(node)
          doc = node.document
          doc.instance_variable_get(:@pptx_wrappers) ||
            doc.instance_variable_set(:@pptx_wrappers, {}.compare_by_identity)
        end

        def element?(object)
          object.is_a?(Nokogiri::XML::Element)
        end

        def local_name(node)
          node.name
        end

        def namespace_uri(node)
          node.namespace&.href
        end

        def parent_element(node)
          parent = node.parent
          parent.is_a?(Nokogiri::XML::Element) ? parent : nil
        end

        def element_children(node)
          node.element_children.to_a
        end

        def xpath(node, expression, namespaces)
          node.xpath(expression, namespaces).to_a
        end

        def append_child(parent, child)
          parent.add_child(child)
        end

        def insert_before(reference, node)
          reference.add_previous_sibling(node)
        end

        def insert_after(reference, node)
          reference.add_next_sibling(node)
        end

        def prepend_child(parent, child)
          parent.prepend_child(child)
        end

        def unlink(node)
          node.unlink
        end

        def text(node)
          node.text
        end

        def set_text(node, value)
          node.content = value
        end

        def get_attribute(node, name)
          node[name]
        end

        def set_attribute(node, name, value)
          node[name] = value
        end

        def remove_attribute(node, name)
          node.remove_attribute(name)
        end

        def get_attribute_ns(node, local, uri)
          node.attribute_with_ns(local, uri)&.value
        end

        # Sets +prefix+:+local+, declaring +prefix+ on +node+ when not in scope.
        def set_attribute_ns(node, prefix, local, uri, value)
          if (existing = node.attribute_with_ns(local, uri))
            existing.value = value
          else
            declare_namespace(node, prefix, uri)
            node["#{prefix}:#{local}"] = value
          end
        end

        def remove_attribute_ns(node, local, uri)
          node.attribute_with_ns(local, uri)&.unlink
        end

        # Reuse an in-scope declaration when there is one, so the output does
        # not sprout a redundant xmlns on every new element.
        def declare_namespace(node, prefix, uri)
          existing = node.namespace_scopes.find { |ns| ns.href == uri && ns.prefix == prefix }
          existing || node.add_namespace_definition(prefix, uri)
        end

        def to_xml(node)
          node.to_xml(save_with: Nokogiri::XML::Node::SaveOptions::AS_XML |
                                 Nokogiri::XML::Node::SaveOptions::NO_DECLARATION)
        end

        def pretty_xml(node)
          node.to_xml(indent: 2)
        end
      end

      # -----------------------------------------------------------------------
      # REXML
      # -----------------------------------------------------------------------
      #
      # REXML resolves a prefix by walking up to the nearest declaration, so a
      # node outside a tree has no namespace at all. A node built, parsed or
      # removed for insertion elsewhere therefore carries its own declarations
      # of the prefixes it uses. Inserting a node drops every declaration in
      # it that repeats one in scope where it lands -- which is what Nokogiri
      # does -- so the output matches Nokogiri's.
      module RexmlBackend
        OWNER = :@pptx_owner
        WRAPPER = :@pptx_wrapper
        FRAGMENT_ROOT = "pptx-fragment"

        module_function

        def name
          :rexml
        end

        def load
          require "rexml/document"
          require "rexml/xpath"
        end

        def parse(xml)
          load
          doc = REXML::Document.new(utf8(xml))
          root = doc.root or raise InvalidXmlError, "no root element"
          strip_blanks(root)
          root
        rescue REXML::ParseException, REXML::UndefinedNamespaceException => e
          raise InvalidXmlError, e.message.lines.first.to_s.strip
        end

        def parse_fragment(context, xml)
          scope = in_scope_namespaces(context)
          decls = scope.map { |prefix, uri| prefix.empty? ? %(xmlns="#{uri}") : %(xmlns:#{prefix}="#{uri}") }
          wrapper = parse(%(<#{FRAGMENT_ROOT} #{decls.join(" ")}>#{utf8(xml)}</#{FRAGMENT_ROOT}>))
          fragment = element_children(wrapper).first
          return nil if fragment.nil?

          # As libxml2 does when parsing in context: a declaration the context
          # already makes is dropped, at any depth; a new namespace or a
          # rebinding of a prefix is kept. The root keeps its own until it is
          # inserted, so the detached fragment still resolves.
          element_children(fragment).each { |child| drop_redundant_declarations(child) }
          fragment.remove
          carry_namespaces(fragment, scope)
          fragment.instance_variable_set(OWNER, document(context))
          fragment
        end

        def build(context, prefix, local, uri)
          scope = in_scope_namespaces(context)
          bound = if scope[prefix] == uri
                    prefix
                  else
                    scope.find { |_, scoped_uri| scoped_uri == uri }&.first
                  end
          child = REXML::Element.new(bound.nil? || bound.empty? ? local : "#{bound}:#{local}")
          declare(child, bound || prefix, uri)
          child.instance_variable_set(OWNER, document(context))
          child
        end

        def document(node)
          top = node
          top = top.parent while top.parent
          top.is_a?(REXML::Document) ? top : top.instance_variable_get(OWNER)
        end

        # A wrapper per node, held on the node itself. REXML nodes keep their
        # identity when moved, so no per-document table is needed.
        def wrapper_cache(_node)
          WrapperSlot
        end

        # Duck-types the Hash#[] / #[]= that {Element.wrap} uses.
        module WrapperSlot
          def self.[](node)
            node.instance_variable_get(WRAPPER)
          end

          def self.[]=(node, wrapper)
            node.instance_variable_set(WRAPPER, wrapper)
          end
        end

        def element?(object)
          object.is_a?(REXML::Element) && !object.is_a?(REXML::Document)
        end

        def local_name(node)
          node.name
        end

        def namespace_uri(node)
          uri = resolve(node, node.prefix.to_s)
          uri.nil? || uri.empty? ? nil : uri
        end

        def parent_element(node)
          parent = node.parent
          element?(parent) ? parent : nil
        end

        def element_children(node)
          node.children.grep(REXML::Element)
        end

        def xpath(node, expression, namespaces)
          REXML::XPath.match(node, expression, namespaces)
        end

        def append_child(parent, child)
          child.remove if child.parent
          parent.add(child)
          settle(child)
        end

        def insert_before(reference, node)
          node.remove if node.parent
          reference.parent.insert_before(reference, node)
          settle(node)
        end

        def insert_after(reference, node)
          node.remove if node.parent
          reference.parent.insert_after(reference, node)
          settle(node)
        end

        def prepend_child(parent, child)
          first = parent.children.first
          return append_child(parent, child) if first.nil?

          insert_before(first, child)
        end

        # A removed node keeps working, as Nokogiri's does: it takes along
        # declarations for the prefixes it used from its old scope, so it can
        # still be read -- or inserted again.
        def unlink(node)
          owner = document(node)
          scope = node.parent ? in_scope_namespaces(node.parent) : {}
          node.remove
          carry_namespaces(node, scope)
          node.instance_variable_set(OWNER, owner)
          node
        end

        def text(node)
          node.children.each_with_object(+"") do |child, out|
            case child
            when REXML::Text then out << child.value
            when REXML::Element then out << text(child)
            end
          end
        end

        def set_text(node, value)
          node.children.dup.each(&:remove)
          node.add_text(REXML::Text.new(value, true, nil, false)) unless value.empty?
        end

        def get_attribute(node, name)
          node.attributes.each_attribute do |attribute|
            return attribute.value if attribute.prefix.empty? && attribute.name == name
          end
          nil
        end

        def set_attribute(node, name, value)
          node.add_attribute(name, value)
        end

        def remove_attribute(node, name)
          node.delete_attribute(name)
        end

        def get_attribute_ns(node, local, uri)
          find_attribute_ns(node, local, uri)&.value
        end

        def set_attribute_ns(node, prefix, local, uri, value)
          if (existing = find_attribute_ns(node, local, uri))
            node.add_attribute(existing.expanded_name, value)
          else
            declare_namespace(node, prefix, uri)
            node.add_attribute("#{prefix}:#{local}", value)
          end
        end

        def remove_attribute_ns(node, local, uri)
          attribute = find_attribute_ns(node, local, uri)
          node.attributes.delete(attribute) if attribute
        end

        def declare_namespace(node, prefix, uri)
          declare(node, prefix, uri) unless resolve(node, prefix) == uri
        end

        def to_xml(node)
          write(node, +"", nil, 0).force_encoding(Encoding::UTF_8)
        end

        def pretty_xml(node)
          write(node, +"", "  ", 0).force_encoding(Encoding::UTF_8)
        end

        # -- internals ---------------------------------------------------------

        def utf8(xml)
          xml = xml.dup if xml.frozen?
          xml.force_encoding(Encoding::UTF_8)
        end

        # libxml2's noblanks, which Nokogiri parses with: whitespace between
        # elements goes, whitespace that is an element's only content stays --
        # `<a:t> </a:t>` is a space, not formatting.
        def strip_blanks(node)
          children = node.children
          if children.any?(REXML::Element) && children.none? { |c| c.is_a?(REXML::Text) && !blank?(c) }
            children.grep(REXML::Text).each(&:remove)
          end
          element_children(node).each { |child| strip_blanks(child) }
        end

        def blank?(text)
          text.to_s.match?(/\A[ \t\r\n]*\z/)
        end

        # The URI +prefix+ ("" for the default namespace) is bound to at
        # +node+, or nil.
        def resolve(node, prefix)
          name = prefix.empty? ? "xmlns" : "xmlns:#{prefix}"
          current = node
          while current.is_a?(REXML::Element) && !current.is_a?(REXML::Document)
            declared = current.attributes.get_attribute(name)
            return declared.value if declared

            current = current.parent
          end
          prefix == "xml" ? "http://www.w3.org/XML/1998/namespace" : nil
        end

        # prefix => URI for every namespace in scope at +node+, nearest first.
        def in_scope_namespaces(node)
          scope = {}
          current = node
          while current.is_a?(REXML::Element) && !current.is_a?(REXML::Document)
            current.attributes.each_attribute do |attribute|
              if attribute.prefix == "xmlns"
                scope[attribute.name] ||= attribute.value
              elsif attribute.expanded_name == "xmlns"
                scope[""] ||= attribute.value
              end
            end
            current = current.parent
          end
          scope
        end

        def in_scope_namespaces_of_self(node)
          node.attributes.each_attribute.with_object({}) do |attribute, decls|
            if attribute.prefix == "xmlns"
              decls[attribute.name] = attribute.value
            elsif attribute.expanded_name == "xmlns"
              decls[""] = attribute.value
            end
          end
        end

        def used_prefixes(node, found = Set.new)
          found << node.prefix.to_s
          node.attributes.each_attribute do |attribute|
            prefix = attribute.prefix.to_s
            found << prefix unless prefix.empty? || prefix == "xmlns"
          end
          element_children(node).each { |child| used_prefixes(child, found) }
          found.to_a
        end

        # Declare on detached +node+ each prefix it uses that +scope+ bound for
        # it and it does not declare itself.
        def carry_namespaces(node, scope)
          added = used_prefixes(node).select { |prefix| scope.key?(prefix) } -
                  in_scope_namespaces_of_self(node).keys
          return if added.empty?

          added.each { |prefix| declare(node, prefix, scope[prefix]) }
        end

        def drop_redundant_declarations(node)
          in_scope_namespaces_of_self(node).each do |prefix, uri|
            next unless resolve(node.parent, prefix) == uri

            node.attributes.delete(node.attributes.get_attribute(prefix.empty? ? "xmlns" : "xmlns:#{prefix}"))
          end
          element_children(node).each { |child| drop_redundant_declarations(child) }
        end

        def declare(node, prefix, uri)
          if prefix.nil? || prefix.empty?
            node.add_attribute("xmlns", uri)
          else
            node.add_attribute("xmlns:#{prefix}", uri)
          end
        end

        # What Nokogiri does on every insertion: drop each declaration in the
        # inserted subtree that repeats one already in scope where it landed.
        def settle(node)
          node.remove_instance_variable(OWNER) if node.instance_variable_defined?(OWNER)
          drop_redundant_declarations(node)
          node
        end

        def find_attribute_ns(node, local, uri)
          node.attributes.each_attribute do |attribute|
            prefix = attribute.prefix.to_s
            next if prefix.empty? || prefix == "xmlns" || attribute.name != local

            return attribute if resolve(node, prefix) == uri
          end
          nil
        end

        # Serializes as libxml2 does with AS_XML: namespace declarations
        # first, double-quoted attributes, empty elements self-closed.
        # +indent+ nil writes compactly; otherwise elements holding only
        # elements are laid out one child per line.
        def write(node, out, indent, depth)
          case node
          when REXML::Element then write_element(node, out, indent, depth)
          when REXML::CData then out << "<![CDATA[" << node.value << "]]>"
          when REXML::Text then out << node.value.gsub(/[&<>\r]/, TEXT_ESCAPES)
          when REXML::Comment then out << "<!--" << node.string << "-->"
          when REXML::Instruction then out << "<?" << node.target << " " << node.content.to_s << "?>"
          end
          out
        end

        def write_element(node, out, indent, depth)
          out << "<" << node.expanded_name
          write_attributes(node, out)
          return out << "/>" if node.children.empty?

          out << ">"
          write_children(node.children, out, indent, depth)
          out << "</" << node.expanded_name << ">"
        end

        def write_attributes(node, out)
          declarations, plain = node.attributes.each_attribute.partition do |a|
            a.prefix == "xmlns" || a.expanded_name == "xmlns"
          end
          (declarations + plain).each do |attribute|
            out << " " << attribute.expanded_name << '="'
            out << attribute.value.gsub(/[&<>"\n\r\t]/, ATTR_ESCAPES) << '"'
          end
        end

        # Mixed content is written as is, as libxml2 does: indenting it would
        # add text the document does not have.
        def write_children(children, out, indent, depth)
          unless indent && children.all?(REXML::Element)
            children.each { |child| write(child, out, nil, depth + 1) }
            return
          end

          children.each do |child|
            out << "\n" << (indent * (depth + 1))
            write(child, out, indent, depth + 1)
          end
          out << "\n" << (indent * depth)
        end
      end
    end
  end
end
