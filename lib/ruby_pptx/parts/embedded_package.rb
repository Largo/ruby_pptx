# frozen_string_literal: true

require "ruby_pptx/opc/package"
require "ruby_pptx/opc/constants"
require "ruby_pptx/enum/prog_id"
require "ruby_pptx/parts/chart"

module Pptx
  module Parts
    # A file embedded in the presentation as an OLE object.
    #
    # An Office document -- one of Enum::PROG_ID -- is stored as the document
    # itself, under a name and content type saying what it is. Anything else
    # is stored as an opaque `oleObjectN.bin`.
    module EmbeddedPackagePart
      OFFICE_PARTS = {
        DOCX: ["/ppt/embeddings/Microsoft_Word_Document%d.docx", Opc::CONTENT_TYPE::WML_DOCUMENT],
        PPTX: ["/ppt/embeddings/Microsoft_PowerPoint_Presentation%d.pptx",
               Opc::CONTENT_TYPE::PML_PRESENTATION],
        XLSX: [EmbeddedXlsxPart::PARTNAME_TEMPLATE, Opc::CONTENT_TYPE::SML_SHEET]
      }.freeze

      GENERIC_PARTNAME = "/ppt/embeddings/oleObject%d.bin"

      module_function

      # @param prog_id [Enum::PROG_ID::Member, String]
      # @return [Opc::Part]
      def new_part(prog_id, blob, package)
        unless prog_id.is_a?(Enum::PROG_ID::Member)
          return Opc::Part.new(package.next_partname(GENERIC_PARTNAME),
                               Opc::CONTENT_TYPE::OFC_OLE_OBJECT, package, blob)
        end

        # An embedded workbook is the same kind of part a chart keeps its data in.
        return EmbeddedXlsxPart.new_xlsx(blob, package) if prog_id.name == :XLSX

        template, content_type = OFFICE_PARTS.fetch(prog_id.name)
        Opc::Part.new(package.next_partname(template), content_type, package, blob)
      end
    end
  end
end
