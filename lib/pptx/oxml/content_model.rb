# frozen_string_literal: true

require "pptx/oxml/element"

module Pptx
  module Oxml
    # The class macros that declare an element's content model.
    #
    # This replaces python-pptx's `MetaOxmlElement` metaclass. Where Python
    # collected descriptor objects out of the class dict and post-processed
    # them, Ruby can just generate the methods as each macro runs, which reads
    # far more directly.
    #
    # Generated method names follow the declared property name +name+:
    #
    # | macro               | generates                                            |
    # |---------------------|------------------------------------------------------|
    # | `optional_attr`     | `name`, `name=`                                      |
    # | `required_attr`     | `name`, `name=`                                      |
    # | `one_and_only_one`  | `name` (raises if absent)                            |
    # | `zero_or_one`       | `name`, `get_or_add_name`, `add_name`, `new_name`,
    #                         `insert_name`, `remove_name` |
    # | `zero_or_more`      | `name_list`, `add_name`, `new_name`, `insert_name`    |
    # | `one_or_more`       | as `zero_or_more`                                    |
    # | `zero_or_one_choice`| `name`, `remove_name`, plus per-choice accessors and `get_or_change_to_x` |
    #
    # python-pptx generates a private `_add_x` and, for repeating elements, a
    # separate public `add_x`. Here they are one `add_x` taking keyword
    # arguments, which is simpler but means a hand-written convenience wrapper
    # cannot reuse the name -- call it something like `add_x_for` instead.
    module ContentModel
      # A member of an `EG_*` element group, for {#zero_or_one_choice}.
      Choice = Data.define(:nsptag) do
        def prop_name = Ns.split_tag(nsptag).last
      end

      # Declare one member of a choice group.
      def choice(nsptag) = Choice.new(nsptag)

      # An attribute that may be absent. Reading returns +default+ when it is;
      # assigning +default+ removes it, so a document never carries an
      # attribute that only restates the schema default.
      def optional_attr(xml_name, type:, default: nil, as: nil)
        name = as || attr_prop_name(xml_name)

        define_method(name) do
          raw = get(xml_name)
          raw.nil? ? default : type.from_xml(raw)
        end

        define_method("#{name}=") do |value|
          if value == default || value.nil?
            delete_attribute(xml_name)
          else
            set(xml_name, type.to_xml(value))
          end
          value
        end
      end

      # An attribute the schema requires. Reading raises when it is missing,
      # because that means the document is invalid rather than merely sparse.
      def required_attr(xml_name, type:, as: nil)
        name = as || attr_prop_name(xml_name)

        define_method(name) do
          raw = get(xml_name)
          if raw.nil?
            raise InvalidXmlError,
                  "required #{xml_name.inspect} attribute not present on <#{nsptag}>"
          end
          type.from_xml(raw)
        end

        define_method("#{name}=") { |value| set(xml_name, type.to_xml(value)) }
      end

      # A required child element.
      def one_and_only_one(nsptag, as: nil)
        name = as || Ns.split_tag(nsptag).last

        define_method(name) do
          find(nsptag) ||
            raise(InvalidXmlError, "required <#{nsptag}> child element not present")
        end
      end

      # An optional child element.
      def zero_or_one(nsptag, successors: [], as: nil)
        name = as || Ns.split_tag(nsptag).last

        define_method(name) { find(nsptag) }
        define_creator(name, nsptag)
        define_inserter(name, successors)
        define_adder(name)
        define_once("get_or_add_#{name}") { find(nsptag) || public_send("add_#{name}") }
        define_once("remove_#{name}") { remove_all(nsptag) }
      end

      # A child element that may repeat. The singular accessor is deliberately
      # not generated -- with a repeating element there is no "the" one.
      def zero_or_more(nsptag, successors: [], as: nil)
        name = as || Ns.split_tag(nsptag).last

        define_method("#{name}_list") { find_all(nsptag) }
        define_creator(name, nsptag)
        define_inserter(name, successors)
        define_adder(name)
      end

      # A child element that must appear at least once. Structurally identical
      # to {#zero_or_more}; the distinction is documentation of the schema.
      def one_or_more(nsptag, successors: [], as: nil)
        zero_or_more(nsptag, successors: successors, as: as)
      end

      # An `EG_*` group, at most one member of which may be present.
      def zero_or_one_choice(choices, as:, successors: [])
        member_tags = choices.map(&:nsptag)

        define_method(as) { first_child_found_in(*member_tags) }
        define_once("remove_#{as}") { remove_all(*member_tags) }

        choices.each { |c| define_choice(c, as, successors) }
      end

      private

      def define_choice(choice, group_name, successors)
        name = choice.prop_name
        nsptag = choice.nsptag

        define_method(name) { find(nsptag) }
        define_creator(name, nsptag)
        define_inserter(name, successors)
        define_adder(name)

        # Swapping in a different group member means clearing whichever one is
        # there now; two members present at once would be invalid.
        define_once("get_or_change_to_#{name}") do
          find(nsptag) || begin
            public_send("remove_#{group_name}")
            public_send("add_#{name}")
          end
        end
      end

      def define_creator(name, nsptag)
        define_once("new_#{name}") { build(nsptag) }
      end

      def define_inserter(name, successors)
        define_once("insert_#{name}") do |child|
          insert_element_before(child, *successors)
        end
      end

      def define_adder(name)
        define_once("add_#{name}") do |**attrs|
          child = public_send("new_#{name}")
          attrs.each { |attr, value| child.public_send("#{attr}=", value) }
          public_send("insert_#{name}", child)
          child
        end
      end

      # Define +name+ unless the class (or an ancestor) already has it, so a
      # hand-written override placed above the macro wins. A definition placed
      # below the macro wins by simply replacing it.
      def define_once(name, &)
        return if method_defined?(name) || private_method_defined?(name)

        define_method(name, &)
      end

      # "r:embed" -> "embed", "macro" -> "macro"
      def attr_prop_name(xml_name)
        xml_name.to_s.include?(":") ? Ns.split_tag(xml_name.to_s).last : xml_name.to_s
      end
    end

    class Element
      extend ContentModel
    end
  end
end
