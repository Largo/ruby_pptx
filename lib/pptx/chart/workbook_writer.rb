# frozen_string_literal: true

require "zip"
require "stringio"

module Pptx
  # Writes the small Excel workbook embedded alongside a chart.
  #
  # PowerPoint opens this workbook when the user chooses "Edit Data"; it is not
  # used for rendering, which is driven entirely by the cached values in the
  # chart XML itself.
  #
  # python-pptx delegates this to XlsxWriter. This writes the minimum a
  # conforming consumer needs directly, which keeps the gem free of a
  # spreadsheet dependency. Strings are written inline rather than through a
  # shared-strings part, which is valid and removes a whole part from the
  # package.
  class ChartWorkbookWriter
    SHEET_NAME = "Sheet1"

    NS_SPREADSHEETML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    NS_OFC_RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    NS_PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"

    CT_SHEET_MAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    RT_OFFICE_DOCUMENT = "#{NS_OFC_RELS}/officeDocument".freeze
    RT_WORKSHEET = "#{NS_OFC_RELS}/worksheet".freeze

    def initialize(chart_data)
      @chart_data = chart_data
    end

    # @return [String] the .xlsx bytes
    def blob
      buffer = Zip::OutputStream.write_buffer(StringIO.new) do |zos|
        entries.each do |name, content|
          zos.put_next_entry(name)
          zos.write(content)
        end
      end
      buffer.rewind
      buffer.read
    end

    private

    def entries
      {
        "[Content_Types].xml" => content_types_xml,
        "_rels/.rels" => package_rels_xml,
        "xl/workbook.xml" => workbook_xml,
        "xl/_rels/workbook.xml.rels" => workbook_rels_xml,
        "xl/worksheets/sheet1.xml" => worksheet_xml
      }
    end

    def declaration = %(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n)

    def content_types_xml
      <<~XML
        #{declaration.chomp}
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="#{CT_SHEET_MAIN}"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="#{CT_WORKSHEET}"/>
        </Types>
      XML
    end

    def package_rels_xml
      <<~XML
        #{declaration.chomp}
        <Relationships xmlns="#{NS_PKG_RELS}">
        <Relationship Id="rId1" Type="#{RT_OFFICE_DOCUMENT}" Target="xl/workbook.xml"/>
        </Relationships>
      XML
    end

    def workbook_xml
      <<~XML
        #{declaration.chomp}
        <workbook xmlns="#{NS_SPREADSHEETML}" xmlns:r="#{NS_OFC_RELS}">
        <sheets><sheet name="#{SHEET_NAME}" sheetId="1" r:id="rId1"/></sheets>
        </workbook>
      XML
    end

    def workbook_rels_xml
      <<~XML
        #{declaration.chomp}
        <Relationships xmlns="#{NS_PKG_RELS}">
        <Relationship Id="rId1" Type="#{RT_WORKSHEET}" Target="worksheets/sheet1.xml"/>
        </Relationships>
      XML
    end

    def worksheet_xml
      "#{declaration}<worksheet xmlns=\"#{NS_SPREADSHEETML}\"><sheetData>" \
        "#{rows_xml}</sheetData></worksheet>\n"
    end

    # Row 1 holds the series names, starting in column B; each later row holds
    # one category label and that category's value in every series.
    def rows_xml
      rows = [header_row_xml]
      @chart_data.category_count.times do |i|
        rows << category_row_xml(i)
      end
      rows.join
    end

    def header_row_xml
      cells = @chart_data.series.map do |series|
        string_cell("#{@chart_data.column_letter(series)}1", series.name)
      end
      %(<row r="1">#{cells.join}</row>)
    end

    def category_row_xml(index)
      row = ChartData::FIRST_DATA_ROW + index
      category = @chart_data.categories[index]
      cells = [cell_for("A#{row}", category)]
      cells += @chart_data.series.map do |series|
        number_cell("#{@chart_data.column_letter(series)}#{row}", series.values[index])
      end
      %(<row r="#{row}">#{cells.join}</row>)
    end

    def cell_for(reference, value)
      value.is_a?(Numeric) ? number_cell(reference, value) : string_cell(reference, value)
    end

    def string_cell(reference, value)
      %(<c r="#{reference}" t="inlineStr"><is><t>#{escape(value)}</t></is></c>)
    end

    # A nil value means "no data for this category", which is written as an
    # empty cell rather than a zero.
    def number_cell(reference, value)
      return %(<c r="#{reference}"/>) if value.nil?

      %(<c r="#{reference}"><v>#{format_number(value)}</v></c>)
    end

    # Integral floats are written without a trailing ".0", as Excel does.
    def format_number(value)
      return value.to_s unless value.is_a?(Float)

      value == value.truncate ? value.to_i.to_s : value.to_s
    end

    def escape(value)
      value.to_s.gsub("&", "&amp;").gsub("<", "&lt;").gsub(">", "&gt;")
    end
  end
end
