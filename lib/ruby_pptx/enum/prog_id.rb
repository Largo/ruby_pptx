# frozen_string_literal: true

module Pptx
  module Enum
    # The Office file types an OLE object can embed as a document of its own,
    # with the icon PowerPoint shows for each.
    #
    # Written by hand rather than generated: unlike the other enumerations,
    # each member carries a ProgID, an icon and a default size, not an MS API
    # integer. An object of any other type is embedded by giving its ProgID
    # as a String, e.g. "Package" or "AcroExch.Document".
    module PROG_ID
      # One embeddable Office file type.
      Member = Data.define(:name, :prog_id, :icon_filename, :width, :height) do
        def to_s = "PROG_ID.#{name}"
        alias_method :inspect, :to_s
      end

      DOCX = Member.new(name: :DOCX, prog_id: "Word.Document.12", icon_filename: "docx-icon.emf",
                        width: 965_200, height: 609_600)
      PPTX = Member.new(name: :PPTX, prog_id: "PowerPoint.Show.12", icon_filename: "pptx-icon.emf",
                        width: 965_200, height: 609_600)
      XLSX = Member.new(name: :XLSX, prog_id: "Excel.Sheet.12", icon_filename: "xlsx-icon.emf",
                        width: 965_200, height: 609_600)

      MEMBERS = [DOCX, PPTX, XLSX].freeze

      # A member by name, e.g. :xlsx; anything else is returned as given, so a
      # ProgID String passes straight through.
      def self.resolve(value)
        return value unless value.is_a?(Symbol)

        MEMBERS.find { |member| member.name == value.upcase } ||
          raise(ArgumentError, "#{value.inspect} is not a member of PROG_ID")
      end
    end
  end
end
