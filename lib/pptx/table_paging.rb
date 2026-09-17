# frozen_string_literal: true

module Pptx
  # Splits a long table across as many slides as it needs.
  #
  # PowerPoint has no notion of a table that flows: a table lives on one slide
  # and is clipped if it outgrows it. Paging it is therefore something the
  # caller has to do, and this does it — repeating the header on each page so
  # every slide stands on its own.
  #
  #   prs.slides.add_table_pages(rows,
  #                              layout: prs.slide_layouts["Blank"],
  #                              left: Pptx.inches(0.5), top: Pptx.inches(1),
  #                              width: Pptx.inches(9), height: Pptx.inches(5))
  module TablePaging
    # The row height assumed when working out how many rows fit, unless the
    # caller says otherwise.
    DEFAULT_ROW_HEIGHT = Length.emu(365_760) # 0.4in

    module_function

    # Split +rows+ into pages of at most +rows_per_page+ body rows, repeating
    # +header+ at the top of each.
    #
    # @return [Array<Array<Array>>] one array of rows per page
    def paginate(rows, rows_per_page:, header: nil)
      body = rows.to_a
      return [] if body.empty?

      raise ArgumentError, "rows_per_page must be positive" unless rows_per_page.positive?

      body.each_slice(rows_per_page).map do |page|
        header ? [header] + page : page
      end
    end

    # How many body rows fit in +height+, leaving room for a repeated header.
    def rows_per_page(height:, row_height: DEFAULT_ROW_HEIGHT, header: false)
      available = Length.coerce(height).emu
      per_row = Length.coerce(row_height).emu
      raise ArgumentError, "row_height must be positive" unless per_row.positive?

      fitting = available / per_row
      fitting -= 1 if header
      [fitting, 1].max
    end
  end

  class Slides
    # Add one slide per page of +rows+, each carrying a table.
    #
    # The first row is treated as a header and repeated on every page unless
    # +header+ is false. When +rows_per_page+ is omitted it is derived from
    # +height+ and +row_height+.
    #
    # @return [Array<Slide>] the slides created, in order
    def add_table_pages(rows, layout:, left:, top:, width:, height:,
                        rows_per_page: nil, header: true,
                        row_height: TablePaging::DEFAULT_ROW_HEIGHT)
      rows = rows.to_a
      return [] if rows.empty?

      header_row = header ? rows.first : nil
      body = header ? rows.drop(1) : rows
      return [] if body.empty?

      per_page = rows_per_page || TablePaging.rows_per_page(
        height: height, row_height: row_height, header: !header_row.nil?
      )

      TablePaging.paginate(body, rows_per_page: per_page, header: header_row).map do |page_rows|
        slide = add(layout)
        fill_table(slide, page_rows, left, top, width, height)
        slide
      end
    end

    private

    def fill_table(slide, page_rows, left, top, width, height)
      columns = page_rows.map(&:size).max
      frame = slide.shapes.add_table(page_rows.size, columns,
                                     at: [left, top], size: [width, height])
      table = frame.table
      page_rows.each_with_index do |row, row_index|
        row.each_with_index { |value, col| table.cell(row_index, col).text = value.to_s }
      end
      table
    end
  end
end
