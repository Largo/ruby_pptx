# frozen_string_literal: true

require "ruby_pptx/errors"

module Pptx
  # Enumerations mirroring those in the Microsoft PowerPoint object model.
  #
  # Each member carries the integer the MS API assigns it, and -- for the
  # enumerations that map to XML -- the attribute value used in the file
  # format. Members compare equal to their MS API integer, so values read from
  # the MS documentation can be used directly.
  module Enum
    # One member of an enumeration.
    class Member
      include Comparable

      attr_reader :enum, :name, :value, :xml_value

      def initialize(enum, name, value, xml_value = nil)
        @enum = enum
        @name = name
        @value = value
        @xml_value = xml_value
        freeze
      end

      # True when this member has no XML representation. Such members are
      # return values only: they describe a state the file format expresses by
      # some other means, so they can be returned but never written.
      def xml_value? = !(@xml_value.nil? || @xml_value.empty?)

      def to_i = @value
      alias to_int to_i

      def to_sym = @name

      def <=>(other)
        case other
        when Member then @value <=> other.value
        when Integer then @value <=> other
        end
      end

      def ==(other) = (self <=> other)&.zero? || false
      alias eql? ==

      def hash = @value.hash

      # e.g. "MIDDLE (3)"
      def to_s = "#{@name} (#{@value})"

      def inspect = "#{@enum.short_name}.#{@name}"
    end

    # Base for every enumeration. Members are declared with {.member} and
    # exposed both as constants and through the collection methods.
    class Base
      extend Enumerable

      class << self
        # Declare a member, defining a constant of the same name.
        #
        # @param name [Symbol]
        # @param value [Integer] the MS API integer for this member
        # @param xml [String, nil] the XML attribute value, for XML-mapped enums
        def member(name, value, xml: nil)
          member = Member.new(self, name, value, xml)
          const_set(name, member)
          members << member
          member
        end

        def members = @members ||= []

        def each(&) = members.each(&)

        def short_name = name.to_s.split("::").last

        # Look a member up by symbolic name or MS API value.
        #
        # @return [Member, nil]
        # Member names are SCREAMING_SNAKE to match the Microsoft API, but a
        # caller writing Ruby should be able to say `:oval` rather than
        # `:OVAL`, so lookup ignores case.
        def [](key)
          case key
          when Member then members.include?(key) ? key : nil
          when Symbol, String then find_by_name(key.to_s)
          when Integer then members.find { |m| m.value == key }
          end
        end

        def find_by_name(name)
          members.find { |m| m.name.to_s == name } ||
            members.find { |m| m.name.to_s.casecmp?(name) }
        end

        # As {.[]}, but raises when there is no such member.
        def fetch(key)
          self[key] || raise(ArgumentError, "#{key.inspect} is not a member of #{short_name}")
        end

        # Raise unless +value+ is assignable to this enumeration.
        def validate(value)
          return value if self[value]

          raise ArgumentError, "#{value.inspect} is not a member of #{short_name}"
        end
      end
    end

    # An enumeration whose members also map to XML attribute values.
    class XmlBase < Base
      class << self
        # The member matching XML attribute value +xml_value+.
        #
        # @raise [ArgumentError] when the value is empty or unmapped
        def from_xml(xml_value)
          member = members.find { |m| m.xml_value? && m.xml_value == xml_value } unless
            xml_value.nil? || xml_value.empty?

          member || raise(ArgumentError,
                          "#{short_name} has no XML mapping for #{xml_value.inspect}")
        end

        # The XML attribute value for +value+, which may be a member or its MS
        # API integer.
        def to_xml(value)
          member = fetch(value)
          unless member.xml_value?
            raise ArgumentError, "#{short_name}.#{member.name} has no XML representation"
          end

          member.xml_value
        end
      end
    end
  end
end
