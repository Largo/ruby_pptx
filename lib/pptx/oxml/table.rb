# frozen_string_literal: true

require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/oxml/simple_types"
require "pptx/oxml/text"
require "pptx/enum/text"

module Pptx
  module Oxml
    # `a:gridCol`, one column of the table grid.
    class CT_TableCol < Element
      tag "a:gridCol"
      required_attr "w", type: SimpleTypes::ST_Coordinate
    end

    # `a:tblGrid`, the column widths of a table.
    class CT_TableGrid < Element
      tag "a:tblGrid"
      zero_or_more "a:gridCol", as: :gridCol

      def add_grid_col(width) = add_gridCol(w: width)
    end

    # `a:tblPr`, whole-table properties such as banding and header rows.
    class CT_TableProperties < Element
      tag "a:tblPr"
      %w[bandRow bandCol firstRow firstCol lastRow lastCol].each do |name|
        optional_attr name, type: SimpleTypes::XsdBoolean, default: false
      end
    end

    # `a:tcPr`, the properties of one cell.
    class CT_TableCellProperties < Element
      tag "a:tcPr"
      FILL_SUCCESSORS = %w[a:headers a:extLst].freeze

      zero_or_one_choice Oxml::FILL_CHOICES.map { |t| choice(t) },
                         successors: FILL_SUCCESSORS, as: :eg_fillProperties
      optional_attr "anchor", type: Enum::MSO_VERTICAL_ANCHOR
      optional_attr "marL", type: SimpleTypes::ST_Coordinate32, default: Pptx::Length.emu(91_440)
      optional_attr "marR", type: SimpleTypes::ST_Coordinate32, default: Pptx::Length.emu(91_440)
      optional_attr "marT", type: SimpleTypes::ST_Coordinate32, default: Pptx::Length.emu(45_720)
      optional_attr "marB", type: SimpleTypes::ST_Coordinate32, default: Pptx::Length.emu(45_720)
    end

    # `a:tc`, one cell of a table.
    class CT_TableCell < Element
      tag "a:tc"
      zero_or_one "a:txBody", successors: %w[a:tcPr a:extLst]
      zero_or_one "a:tcPr", successors: %w[a:extLst]
      optional_attr "gridSpan", type: SimpleTypes::XsdInt, default: 1
      optional_attr "rowSpan", type: SimpleTypes::XsdInt, default: 1
      optional_attr "hMerge", type: SimpleTypes::XsdBoolean, default: false
      optional_attr "vMerge", type: SimpleTypes::XsdBoolean, default: false

      CELL_XML = <<~XML.freeze
        <a:tc #{Ns.nsdecls("a")}>
          <a:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p/>
          </a:txBody>
          <a:tcPr/>
        </a:tc>
      XML

      # A cell always has a text body, so a new one is created with it rather
      # than empty.
      def self.new_cell(context) = context.build_from_xml(CELL_XML)

      # The `a:tr` this cell sits in.
      def tr = parent

      # The `a:tbl` this cell belongs to.
      def tbl = tr.parent

      def col_idx = tr.tc_list.index(self)

      def row_idx = tbl.tr_list.index(tr)

      # True when this is the top-left cell of a merged range.
      #
      # Only the origin carries the full span; the cells it covers are marked
      # hMerge or vMerge and keep spans of 1.
      def merge_origin?
        return true if gridSpan > 1 && !vMerge

        rowSpan > 1 && !hMerge
      end

      # True when this cell is covered by a merge rather than being its origin.
      def spanned? = hMerge || vMerge

      # Take the paragraphs from +other+ into this cell's text body.
      #
      # An empty source contributes nothing; a single empty paragraph in the
      # target is replaced rather than appended to; and the source is left with
      # one empty paragraph, since the schema requires at least one.
      def append_paragraphs_from(other)
        source = other.get_or_add_txBody
        target = get_or_add_txBody
        return self if source.empty?

        target.clear_content if target.empty?
        source.p_list.each { |paragraph| target.append(paragraph) }
        source.unclear_content
        self
      end
    end

    # `a:tr`, one row of a table.
    class CT_TableRow < Element
      tag "a:tr"
      zero_or_more "a:tc", successors: %w[a:extLst], as: :tc
      required_attr "h", type: SimpleTypes::ST_Coordinate

      def new_tc = CT_TableCell.new_cell(self)

      def add_cell = add_tc
    end

    # `a:tbl`, a table.
    class CT_Table < Element
      tag "a:tbl"
      zero_or_one "a:tblPr", successors: %w[a:tblGrid a:tr]
      one_and_only_one "a:tblGrid"
      zero_or_more "a:tr", successors: [], as: :tr

      # The style PowerPoint applies to a table created through the UI.
      DEFAULT_TABLE_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

      def self.tbl_xml(table_style_id)
        <<~XML
          <a:tbl #{Ns.nsdecls("a")}>
            <a:tblPr firstRow="1" bandRow="1">
              <a:tableStyleId>#{table_style_id}</a:tableStyleId>
            </a:tblPr>
            <a:tblGrid/>
          </a:tbl>
        XML
      end

      # A new `a:tbl` of +rows+ by +cols+ filling the given width and height.
      #
      # The width is divided evenly between columns and the height between
      # rows; the last of each absorbs the rounding remainder so the totals
      # come out exactly.
      def self.new_tbl(context, rows, cols, width, height, table_style_id = DEFAULT_TABLE_STYLE_ID)
        width = Pptx::Length.coerce(width).emu
        height = Pptx::Length.coerce(height).emu
        tbl = context.build_from_xml(tbl_xml(table_style_id))

        col_width = width / cols
        cols.times do |col|
          this_width = col == cols - 1 ? width - ((cols - 1) * col_width) : col_width
          tbl.tblGrid.add_grid_col(Pptx::Length.emu(this_width))
        end

        row_height = height / rows
        rows.times do |row|
          this_height = row == rows - 1 ? height - ((rows - 1) * row_height) : row_height
          tr = tbl.add_tr(h: Pptx::Length.emu(this_height))
          cols.times { tr.add_cell }
        end

        tbl
      end

      # The cell at +row_idx+, +col_idx+.
      def tc(row_idx, col_idx) = tr_list[row_idx].tc_list[col_idx]

      # The rectangle of cells spanned by the two opposite corners +a+ and +b+.
      #
      # The corners may be given in either order and along either diagonal, so
      # the extents are normalised here.
      def cell_range(a, b)
        top, bottom = [a.row_idx, b.row_idx].minmax
        left, right = [a.col_idx, b.col_idx].minmax
        CellRange.new(self, top, left, bottom, right)
      end

      %w[firstRow firstCol lastRow lastCol bandRow bandCol].each do |name|
        define_method(name) { tblPr ? tblPr.public_send(name) : false }

        define_method("#{name}=") do |value|
          get_or_add_tblPr.public_send("#{name}=", value)
          value
        end
      end
    end

    # A rectangular block of `a:tc` elements, used when merging and splitting.
    #
    # It assumes the table's structure does not change while it is alive, so
    # create one, use it, and let it go.
    class CellRange
      attr_reader :row_count, :column_count

      def initialize(tbl, top, left, bottom, right)
        @tbl = tbl
        @top = top
        @left = left
        @bottom = bottom
        @right = right
        @row_count = bottom - top + 1
        @column_count = right - left + 1
      end

      # The range a merge-origin cell already covers.
      def self.from_merge_origin(tc)
        tc.tbl.cell_range(tc, tc.tbl.tc(tc.row_idx + tc.rowSpan - 1,
                                        tc.col_idx + tc.gridSpan - 1))
      end

      # Every cell in the range, left to right then top to bottom.
      def cells
        (@top..@bottom).flat_map { |row| (@left..@right).map { |col| @tbl.tc(row, col) } }
      end

      def top_row_cells = (@left..@right).map { |col| @tbl.tc(@top, col) }

      def left_column_cells = (@top..@bottom).map { |row| @tbl.tc(row, @left) }

      def cells_except_left_column
        (@top..@bottom).flat_map { |row| ((@left + 1)..@right).map { |col| @tbl.tc(row, col) } }
      end

      def cells_except_top_row
        ((@top + 1)..@bottom).flat_map { |row| (@left..@right).map { |col| @tbl.tc(row, col) } }
      end

      # True when any cell in the range is already part of a merge. Merging
      # over an existing merge would produce a table PowerPoint cannot lay out.
      def contains_merged_cell?
        cells.any? { |tc| tc.gridSpan > 1 || tc.rowSpan > 1 || tc.hMerge || tc.vMerge }
      end

      # Gather the text of the whole range into its top-left cell, which is
      # the one that stays visible.
      def move_content_to_origin
        all = cells
        origin = all.first
        all.drop(1).each { |spanned| origin.append_paragraphs_from(spanned) }
        origin
      end
    end
  end
end
