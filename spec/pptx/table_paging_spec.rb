# frozen_string_literal: true

RSpec.describe Pptx::TablePaging do
  describe ".paginate" do
    let(:rows) { (1..7).map { |i| ["Row #{i}", i] } }

    it "splits rows into pages of the given size" do
      pages = described_class.paginate(rows, rows_per_page: 3)
      aggregate_failures do
        expect(pages.size).to eq(3)
        expect(pages.map(&:size)).to eq([3, 3, 1])
        expect(pages.flatten(1)).to eq(rows)
      end
    end

    # Each page has to stand on its own, so the header goes on every one.
    it "repeats the header on every page" do
      header = %w[Name Value]
      pages = described_class.paginate(rows, rows_per_page: 3, header: header)
      aggregate_failures do
        expect(pages.map(&:first)).to all(eq(header))
        expect(pages.map(&:size)).to eq([4, 4, 2])
      end
    end

    it "returns no pages for no rows" do
      expect(described_class.paginate([], rows_per_page: 5)).to eq([])
    end

    it "rejects a non-positive page size" do
      expect { described_class.paginate(rows, rows_per_page: 0) }
        .to raise_error(ArgumentError, /must be positive/)
    end
  end

  describe ".rows_per_page" do
    it "derives a count from the available height" do
      expect(described_class.rows_per_page(height: Pptx.inches(4),
                                           row_height: Pptx.inches(0.5))).to eq(8)
    end

    it "leaves room for a repeated header" do
      expect(described_class.rows_per_page(height: Pptx.inches(4),
                                           row_height: Pptx.inches(0.5),
                                           header: true)).to eq(7)
    end

    # Better one row per slide than an infinite loop or an empty page.
    it "never returns less than one row" do
      expect(described_class.rows_per_page(height: Pptx.inches(0.1),
                                           row_height: Pptx.inches(1),
                                           header: true)).to eq(1)
    end

    it "rejects a non-positive row height" do
      expect { described_class.rows_per_page(height: Pptx.inches(4), row_height: 0) }
        .to raise_error(ArgumentError, /must be positive/)
    end
  end
end

RSpec.describe "Pptx::Slides#add_table_pages" do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:layout) { presentation.slide_layouts["Blank"] }
  let(:header) { ["Region", "Q1", "Q2"] }
  let(:rows) { [header] + (1..17).map { |i| ["Row #{i}", i * 10, i * 20] } }

  def add_pages(**options)
    presentation.slides.add_table_pages(
      rows, layout: layout,
      left: Pptx.inches(0.5), top: Pptx.inches(1),
      width: Pptx.inches(9), height: Pptx.inches(5), **options
    )
  end

  it "adds as many slides as the data needs" do
    slides = add_pages
    aggregate_failures do
      expect(slides.size).to eq(2)
      expect(presentation.slides.size).to eq(2)
    end
  end

  it "repeats the header and splits the body" do
    slides = add_pages
    tables = slides.map { |slide| slide.shapes.first.table }
    aggregate_failures do
      expect(tables.map { |t| t.cell(0, 0).text }).to all(eq("Region"))
      expect(tables.first.row_count).to eq(12)  # header + 11 body rows
      expect(tables.last.row_count).to eq(7)    # header + 6 body rows
      expect(tables.last.cell(6, 0).text).to eq("Row 17")
    end
  end

  it "honours an explicit rows_per_page" do
    slides = add_pages(rows_per_page: 5)
    expect(slides.size).to eq(4)
  end

  it "treats every row as data when header is false" do
    slides = add_pages(header: false, rows_per_page: 6)
    tables = slides.map { |slide| slide.shapes.first.table }
    aggregate_failures do
      expect(slides.size).to eq(3)
      expect(tables.first.cell(0, 0).text).to eq("Region")
      expect(tables[1].cell(0, 0).text).to eq("Row 6")
    end
  end

  it "fits on one slide when the data is short" do
    presentation.slides.add_table_pages(
      [header, ["A", 1, 2]], layout: layout,
      left: 0, top: 0, width: Pptx.inches(9), height: Pptx.inches(5)
    )
    expect(presentation.slides.size).to eq(1)
  end

  it "adds nothing for an empty table or a header with no body" do
    aggregate_failures do
      expect(presentation.slides.add_table_pages([], layout: layout, left: 0, top: 0,
                                                     width: 100, height: 100)).to eq([])
      expect(presentation.slides.add_table_pages([header], layout: layout, left: 0, top: 0,
                                                           width: 100, height: 100)).to eq([])
      expect(presentation.slides.size).to eq(0)
    end
  end

  it "sizes the table to the widest row" do
    ragged = [%w[A B C], %w[1 2], %w[3 4 5 6]]
    slides = presentation.slides.add_table_pages(
      ragged, layout: layout, left: 0, top: 0,
      width: Pptx.inches(9), height: Pptx.inches(5), header: false
    )
    expect(slides.first.shapes.first.table.column_count).to eq(4)
  end
end
