# frozen_string_literal: true

require "pptx/element_proxy"
require "pptx/text/text"
require "pptx/dml/fill"

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

    def part = @parent.part

    # The cell at +row_idx+, +col_idx+.
    def cell(row_idx, col_idx) = Cell.new(@element.tc(row_idx, col_idx), self)

    # `table[0, 1]` reads better than `table.cell(0, 1)` in a loop.
    alias [] cell

    def rows = @rows ||= TableRows.new(@element, self)

    def columns = @columns ||= TableColumns.new(@element, self)

    def row_count = @element.tr_list.size

    def column_count = @element.tblGrid.gridCol_list.size

    # Whether the first row is styled as a header. The same pattern applies to
    # first_col, last_row, last_col, horz_banding and vert_banding.
    def first_row? = @element.firstRow

    def first_row=(value)
      @element.firstRow = value
      value
    end

    def first_col? = @element.firstCol

    def first_col=(value)
      @element.firstCol = value
      value
    end

    def last_row? = @element.lastRow

    def last_row=(value)
      @element.lastRow = value
      value
    end

    def last_col? = @element.lastCol

    def last_col=(value)
      @element.lastCol = value
      value
    end

    # Whether alternate rows are shaded. Spelled out where python-pptx says
    # `horz_banding`.
    def banded_rows? = @element.bandRow

    def banded_rows=(value)
      @element.bandRow = value
      value
    end

    def banded_columns? = @element.bandCol

    def banded_columns=(value)
      @element.bandCol = value
      value
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

    def inspect = "#<Pptx::Table #{row_count}x#{column_count}>"
  end

  # The rows of a table.
  class TableRows
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

    def size = @element.tr_list.size
    alias length size

    def [](index)
      tr = @element.tr_list[index]
      tr && TableRow.new(tr, @table)
    end
  end

  # The columns of a table.
  class TableColumns
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

    def size = @element.tblGrid.gridCol_list.size
    alias length size

    def [](index)
      col = @element.tblGrid.gridCol_list[index]
      col && TableColumn.new(col, @table)
    end
  end

  # One row of a table.
  class TableRow < ElementProxy
    def initialize(tr, table)
      super(tr)
      @table = table
    end

    def height = @element.h

    def height=(value)
      @element.h = value
      @table.notify_height_changed
      value
    end

    def cells = @element.tc_list.map { |tc| Cell.new(tc, @table) }

    def inspect = "#<Pptx::TableRow height=#{height&.inches}in>"
  end

  # One column of a table.
  class TableColumn < ElementProxy
    def initialize(grid_col, table)
      super(grid_col)
      @table = table
    end

    def width = @element.w

    def width=(value)
      @element.w = value
      @table.notify_width_changed
      value
    end

    def inspect = "#<Pptx::TableColumn width=#{width&.inches}in>"
  end

  # One cell of a table.
  class Cell
    attr_reader :element, :parent

    def initialize(tc, parent)
      @element = tc
      @parent = parent
    end

    def part = @parent.part

    def text_frame = @text_frame ||= TextFrame.new(@element.get_or_add_txBody, self)

    def text = text_frame.text

    def text=(value)
      text_frame.text = value
      value
    end

    def fill = @fill ||= FillFormat.from_fill_parent(cell_properties)

    # @return [Pptx::Enum::MSO_ANCHOR, nil]
    def vertical_anchor = cell_properties.anchor

    def vertical_anchor=(value)
      cell_properties.anchor = value
      value
    end

    def margin_left = cell_properties.marL

    def margin_left=(value)
      cell_properties.marL = value
      value
    end

    def margin_right = cell_properties.marR

    def margin_right=(value)
      cell_properties.marR = value
      value
    end

    def margin_top = cell_properties.marT

    def margin_top=(value)
      cell_properties.marT = value
      value
    end

    def margin_bottom = cell_properties.marB

    def margin_bottom=(value)
      cell_properties.marB = value
      value
    end

    # How many columns this cell spans; 1 unless merged.
    def span_width = @element.gridSpan

    # How many rows this cell spans; 1 unless merged.
    def span_height = @element.rowSpan

    # True when this cell is covered by a merged neighbour rather than being
    # the origin of the merge.
    def spanned? = @element.hMerge || @element.vMerge

    def inspect = "#<Pptx::Cell #{text.inspect}>"

    private

    def cell_properties = @element.get_or_add_tcPr
  end
end
