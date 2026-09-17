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

      CELL_XML = <<~XML
        <a:tc #{Ns.nsdecls('a')}>
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
          <a:tbl #{Ns.nsdecls('a')}>
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

      %w[firstRow firstCol lastRow lastCol bandRow bandCol].each do |name|
        define_method(name) { tblPr ? tblPr.public_send(name) : false }

        define_method("#{name}=") do |value|
          get_or_add_tblPr.public_send("#{name}=", value)
          value
        end
      end
    end
  end
end
