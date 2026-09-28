# frozen_string_literal: true

require "ruby_pptx/pattern_matching"

require "ruby_pptx/sliceable"

require "ruby_pptx/element_proxy"
require "ruby_pptx/text/text"
require "ruby_pptx/dml/fill"

module Pptx
  # A table inside a graphic frame.
  #
  #   table = shape.table
  #   table.cell(0, 0).text = "Region"
  #   table.columns[0].width = Pptx.inches(2)
  class Table
    attr_reader :element, :parent

    def initialize(tbl, parent)
      @element = tbl
      @parent = parent
    end

    def part
      @parent.part
    end

    # The cell at +row_idx+, +col_idx+.
    def cell(row_idx, col_idx)
      Cell.new(@element.tc(row_idx, col_idx), self)
    end

    # `table[0, 1]` reads better than `table.cell(0, 1)` in a loop.
    alias [] cell

    def rows
      @rows ||= TableRows.new(@element, self)
    end

    def columns
      @columns ||= TableColumns.new(@element, self)
    end

    def row_count
      @element.tr_list.size
    end

    def column_count
      @element.tblGrid.gridCol_list.size
    end

    # Whether the first row is styled as a header. The same pattern applies to
    # first_col, last_row, last_col, horz_banding and vert_banding.
    def first_row?
      @element.firstRow
    end

    def first_row=(value)
      @element.firstRow = value
    end

    def first_col?
      @element.firstCol
    end

    def first_col=(value)
      @element.firstCol = value
    end

    def last_row?
      @element.lastRow
    end

    def last_row=(value)
      @element.lastRow = value
    end

    def last_col?
      @element.lastCol
    end

    def last_col=(value)
      @element.lastCol = value
    end

    # Whether alternate rows are shaded. Spelled out where python-pptx says
    # `horz_banding`.
    def banded_rows?
      @element.bandRow
    end

    def banded_rows=(value)
      @element.bandRow = value
    end

    def banded_columns?
      @element.bandCol
    end

    def banded_columns=(value)
      @element.bandCol = value
    end

    # @api private
    # Resize the containing graphic frame to the sum of the row heights.
    #
    # A row calls this when its height changes: the frame and the table have
    # to agree on the total, or PowerPoint shows the table clipped or adrift
    # inside its frame.
    def notify_height_changed
      @parent.height = Length.emu(rows.sum { |row| row.height.emu })
    end

    # @api private
    # As {#notify_height_changed}, for column widths.
    def notify_width_changed
      @parent.width = Length.emu(columns.sum { |column| column.width.emu })
    end

    # Iterate every cell, row by row.
    def each_cell
      return enum_for(:each_cell) { row_count * column_count } unless block_given?

      row_count.times { |r| column_count.times { |c| yield cell(r, c) } }
      self
    end

    def inspect
      "#<Pptx::Table #{row_count}x#{column_count}>"
    end
  end

  # The rows of a table.
  class TableRows
    include DeconstructToArray

    include Sliceable
    include Enumerable

    def initialize(tbl, table)
      @element = tbl
      @table = table
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @element.tr_list.each { |tr| yield TableRow.new(tr, @table) }
      self
    end

    def size
      @element.tr_list.size
    end
    alias length size

    def [](index, length = nil)
      slice_members(@element.tr_list, index, length) { |tr| TableRow.new(tr, @table) }
    end
  end

  # The columns of a table.
  class TableColumns
    include DeconstructToArray

    include Sliceable
    include Enumerable

    def initialize(tbl, table)
      @element = tbl
      @table = table
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @element.tblGrid.gridCol_list.each { |col| yield TableColumn.new(col, @table) }
      self
    end

    def size
      @element.tblGrid.gridCol_list.size
    end
    alias length size

    def [](index, length = nil)
      slice_members(@element.tblGrid.gridCol_list, index, length) do |col|
        TableColumn.new(col, @table)
      end
    end
  end

  # One row of a table.
  class TableRow < ElementProxy
    def initialize(tr, table)
      super(tr)
      @table = table
    end

    def height
      @element.h
    end

    def height=(value)
      @element.h = value
      @table.notify_height_changed
    end

    def cells
      @element.tc_list.map { |tc| Cell.new(tc, @table) }
    end

    def inspect
      "#<Pptx::TableRow height=#{height&.inches}in>"
    end
  end

  # One column of a table.
  class TableColumn < ElementProxy
    def initialize(grid_col, table)
      super(grid_col)
      @table = table
    end

    def width
      @element.w
    end

    def width=(value)
      @element.w = value
      @table.notify_width_changed
    end

    def inspect
      "#<Pptx::TableColumn width=#{width&.inches}in>"
    end
  end

  # One cell of a table.
  class Cell
    include PatternMatching

    pattern_keys :text, :span_width, :span_height, :merge_origin?, :spanned?, :vertical_anchor

    attr_reader :element, :parent

    def initialize(tc, parent)
      @element = tc
      @parent = parent
    end

    def part
      @parent.part
    end

    def text_frame
      @text_frame ||= TextFrame.new(@element.get_or_add_txBody, self)
    end

    def text
      text_frame.text
    end

    def text=(value)
      text_frame.text = value
    end

    def fill
      @fill ||= FillFormat.from_fill_parent(cell_properties)
    end

    # @return [Pptx::Enum::MSO_ANCHOR, nil]
    def vertical_anchor
      cell_properties.anchor
    end

    def vertical_anchor=(value)
      cell_properties.anchor = value
    end

    def margin_left
      cell_properties.marL
    end

    def margin_left=(value)
      cell_properties.marL = value
    end

    def margin_right
      cell_properties.marR
    end

    def margin_right=(value)
      cell_properties.marR = value
    end

    def margin_top
      cell_properties.marT
    end

    def margin_top=(value)
      cell_properties.marT = value
    end

    def margin_bottom
      cell_properties.marB
    end

    def margin_bottom=(value)
      cell_properties.marB = value
    end

    # How many columns this cell spans.
    #
    # Only a merge origin carries the real span; on any other cell this reads
    # 1 whether or not it is part of a merge. Test {#merge_origin?} first.
    def span_width
      @element.gridSpan
    end

    # How many rows this cell spans. See {#span_width} on when to trust it.
    def span_height
      @element.rowSpan
    end

    # True when this cell is the top-left of a merged range.
    def merge_origin?
      @element.merge_origin?
    end

    # True when this cell is covered by a merge rather than being its origin.
    def spanned?
      @element.spanned?
    end

    # Merge this cell with +other+, which is the opposite corner of the range.
    #
    # Either diagonal may be given, in either order. The text of every cell in
    # the range is gathered into the top-left one, which is the cell that stays
    # visible.
    #
    # @raise [Error] when the cells are in different tables, or the range
    #   already contains a merge -- merging over a merge produces a table
    #   PowerPoint cannot lay out
    def merge(other)
      raise Error, "cannot merge cells from different tables" unless @element.tbl == other.element.tbl

      range = @element.tbl.cell_range(@element, other.element)
      raise Error, "the range already contains a merged cell" if range.contains_merged_cell?

      range.move_content_to_origin
      range.top_row_cells.each { |tc| tc.rowSpan = range.row_count }
      range.left_column_cells.each { |tc| tc.gridSpan = range.column_count }
      range.cells_except_left_column.each { |tc| tc.hMerge = true }
      range.cells_except_top_row.each { |tc| tc.vMerge = true }
      self
    end

    # Undo a merge, giving back a separate cell for each grid position it
    # covered. The text stays in the origin cell.
    #
    # @raise [Error] unless this is a merge origin
    def split
      raise Error, "only a merge-origin cell can be split" unless merge_origin?

      Oxml::CellRange.from_merge_origin(@element).cells.each do |tc|
        tc.rowSpan = 1
        tc.gridSpan = 1
        tc.hMerge = false
        tc.vMerge = false
      end
      self
    end

    def inspect
      "#<Pptx::Cell #{text.inspect}>"
    end

    private

    def cell_properties
      @element.get_or_add_tcPr
    end
  end
end
