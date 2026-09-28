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
    RT_STYLES = "#{NS_OFC_RELS}/styles".freeze
    CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    # The first id free for a custom number format; 0-163 are built in.
    DATE_FORMAT_ID = 164

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
      parts = {
        "[Content_Types].xml" => content_types_xml,
        "_rels/.rels" => package_rels_xml,
        "xl/workbook.xml" => workbook_xml,
        "xl/_rels/workbook.xml.rels" => workbook_rels_xml,
        "xl/worksheets/sheet1.xml" => worksheet_xml
      }
      parts["xl/styles.xml"] = styles_xml if dates?
      parts
    end

    # Only date categories need a cell style -- a number format that makes
    # the serial numbers read as dates -- so only they get a styles part.
    # XY and bubble data have no categories, so never dates.
    def dates? = @chart_data.respond_to?(:categories) && @chart_data.categories.dates?

    def declaration = %(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n)

    def content_types_xml
      <<~XML
        #{declaration.chomp}
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="#{CT_SHEET_MAIN}"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="#{CT_WORKSHEET}"/>
        #{%(<Override PartName="/xl/styles.xml" ContentType="#{CT_STYLES}"/>) if dates?}
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
        #{%(<Relationship Id="rId2" Type="#{RT_STYLES}" Target="styles.xml"/>) if dates?}
        </Relationships>
      XML
    end

    def worksheet_xml
      "#{declaration}<worksheet xmlns=\"#{NS_SPREADSHEETML}\"><sheetData>" \
        "#{rows_xml}</sheetData></worksheet>\n"
    end

    # Row 1 holds the series names, after the category columns. Each later
    # row holds that row's category labels and one value per series.
    #
    # Grouped categories take one column per level, the leaf level
    # rightmost; a group's label sits on the row where its run begins, as
    # python-pptx writes it.
    def rows_xml
      rows = Hash.new { |hash, row| hash[row] = [] }
      @chart_data.series.each do |series|
        rows[1] << [series.index + depth, string_cell("#{@chart_data.column_letter(series)}1", series.name)]
      end
      category_cells(rows)
      series_cells(rows)
      rows.keys.sort.map do |row|
        %(<row r="#{row}">#{rows[row].sort_by(&:first).map(&:last).join}</row>)
      end.join
    end

    def depth = @chart_data.categories.depth

    def category_cells(rows)
      @chart_data.categories.levels.each_with_index do |level, level_index|
        column = depth - level_index - 1
        letter = ChartData.column_reference(column + 1)
        level.each do |offset, label|
          row = ChartData::FIRST_DATA_ROW + offset
          rows[row] << [column, category_cell("#{letter}#{row}", label)]
        end
      end
    end

    def series_cells(rows)
      @chart_data.series.each do |series|
        letter = @chart_data.column_letter(series)
        series.values.each_with_index do |value, index|
          row = ChartData::FIRST_DATA_ROW + index
          rows[row] << [series.index + depth, number_cell("#{letter}#{row}", value)]
        end
      end
    end

    def category_cell(reference, label)
      return date_cell(reference, label) if label.is_a?(Date) || label.is_a?(Time)

      cell_for(reference, label)
    end

    # A date is its Excel serial number, styled with the date format.
    def date_cell(reference, date)
      category = ChartDataCategory.new(date, nil)
      %(<c r="#{reference}" s="1"><v>#{category.numeric_str_val.to_f.to_i}</v></c>)
    end

    def styles_xml
      format_code = escape(@chart_data.categories.number_format)
      <<~XML
        #{declaration.chomp}
        <styleSheet xmlns="#{NS_SPREADSHEETML}">
        <numFmts count="1"><numFmt numFmtId="#{DATE_FORMAT_ID}" formatCode="#{format_code}"/></numFmts>
        <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
        <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="#{DATE_FORMAT_ID}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
        </styleSheet>
      XML
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
