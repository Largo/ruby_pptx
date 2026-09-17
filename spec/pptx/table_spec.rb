# frozen_string_literal: true

RSpec.describe Pptx::Table do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts[6]) }
  let(:frame) do
    slide.shapes.add_table(3, 2, Pptx.inches(1), Pptx.inches(1),
                           Pptx.inches(6), Pptx.inches(3))
  end
  subject(:table) { frame.table }

  it "is reached through a graphic frame, not directly" do
    aggregate_failures do
      expect(frame).to be_a(Pptx::GraphicFrame)
      expect(frame).to be_table
      expect(frame).not_to be_chart
      expect(frame.shape_type).to eq(Pptx::Enum::MSO_SHAPE_TYPE::TABLE)
    end
  end

  it "raises when a frame is asked for a table it does not hold" do
    graphic_frame = Pptx::GraphicFrame.new(frame.element, slide.shapes)
    allow(frame.element).to receive(:table?).and_return(false)
    expect { graphic_frame.table }.to raise_error(Pptx::Error, /does not contain a table/)
  end

  it "has the requested shape" do
    aggregate_failures do
      expect(table.row_count).to eq(3)
      expect(table.column_count).to eq(2)
      expect(table.each_cell.count).to eq(6)
    end
  end

  it "divides width and height evenly" do
    aggregate_failures do
      expect(table.columns.map(&:width)).to all(eq(Pptx.inches(3)))
      expect(table.rows.map(&:height)).to all(eq(Pptx.inches(1)))
    end
  end

  # An odd total cannot divide evenly, and the remainder has to go somewhere
  # or the table would not fill the frame.
  it "gives the rounding remainder to the last row and column" do
    odd = slide.shapes.add_table(3, 3, 0, 0, 100, 100).table
    aggregate_failures do
      expect(odd.columns.map { |c| c.width.emu }).to eq([33, 33, 34])
      expect(odd.rows.map { |r| r.height.emu }).to eq([33, 33, 34])
      expect(odd.columns.sum { |c| c.width.emu }).to eq(100)
    end
  end

  it "reads and writes cell text" do
    table.cell(0, 0).text = "Region"
    table.cell(2, 1).text = "42"
    aggregate_failures do
      expect(table.cell(0, 0).text).to eq("Region")
      expect(table.cell(2, 1).text).to eq("42")
      expect(table.each_cell.map(&:text)).to eq(["Region", "", "", "", "", "42"])
    end
  end

  it "exposes a cell's text frame for formatting" do
    cell = table.cell(0, 0)
    cell.text = "Header"
    cell.text_frame.paragraphs.first.runs.first.font.bold = true
    expect(cell.element.xml).to include('b="1"')
  end

  it "round-trips cell fill, anchor and margins" do
    cell = table.cell(0, 0)
    cell.fill.solid
    cell.fill.fore_color.rgb = Pptx::RGBColor["1F497D"]
    cell.vertical_anchor = Pptx::Enum::MSO_ANCHOR::MIDDLE
    cell.margin_left = Pptx.inches(0.2)
    aggregate_failures do
      expect(cell.fill.type).to eq(Pptx::Enum::MSO_FILL_TYPE::SOLID)
      expect(cell.vertical_anchor).to eq(Pptx::Enum::MSO_ANCHOR::MIDDLE)
      expect(cell.margin_left).to eq(Pptx.inches(0.2))
    end
  end

  it "defaults to a banded table with a header row, as PowerPoint does" do
    aggregate_failures do
      expect(table.first_row).to be(true)
      expect(table.horz_banding).to be(true)
      expect(table.first_col).to be(false)
      expect(table.last_row).to be(false)
    end
  end

  it "round-trips the style flags" do
    table.first_col = true
    table.horz_banding = false
    aggregate_failures do
      expect(table.first_col).to be(true)
      expect(table.horz_banding).to be(false)
      expect(table.element.xml).to include('firstCol="1"')
    end
  end

  it "resizes rows and columns" do
    table.columns[0].width = Pptx.inches(4)
    table.rows[0].height = Pptx.inches(2)
    aggregate_failures do
      expect(table.columns[0].width).to eq(Pptx.inches(4))
      expect(table.rows[0].height).to eq(Pptx.inches(2))
    end
  end

  # The frame and the table have to agree on the total size, or PowerPoint
  # shows the table clipped or adrift inside its frame.
  it "resizes the graphic frame when a column or row is resized" do
    table.columns[0].width = Pptx.inches(4)
    expect(frame.width).to eq(Pptx.inches(7)) # 4 + 3

    table.rows[0].height = Pptx.inches(2)
    expect(frame.height).to eq(Pptx.inches(4)) # 2 + 1 + 1
  end

  it "reports span and merge state for an unmerged cell" do
    cell = table.cell(0, 0)
    aggregate_failures do
      expect(cell.span_width).to eq(1)
      expect(cell.span_height).to eq(1)
      expect(cell).not_to be_spanned
    end
  end

  it "lists cells row by row" do
    expect(table.rows[0].cells.size).to eq(2)
  end
end

RSpec.describe Pptx::Picture do
  def self.fixture_dir = File.expand_path("../fixtures/images", __dir__)
  def image(name) = File.join(self.class.fixture_dir, name)

  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts[6]) }

  it "uses the image's native size when no size is given" do
    pic = slide.shapes.add_picture(image("png-96dpi.png"), Pptx.inches(1), Pptx.inches(1))
    aggregate_failures do
      expect(pic.width.inches).to be_within(1e-9).of(64 / 96.0)
      expect(pic.height.inches).to be_within(1e-9).of(48 / 96.0)
      expect(pic.shape_type).to eq(Pptx::Enum::MSO_SHAPE_TYPE::PICTURE)
      expect(pic.name).to eq("Picture 1")
    end
  end

  it "preserves the aspect ratio when only one dimension is given" do
    aggregate_failures do
      by_width = slide.shapes.add_picture(image("png-96dpi.png"), 0, 0, width: Pptx.inches(2))
      expect(by_width.height).to eq(Pptx.inches(1.5))

      by_height = slide.shapes.add_picture(image("png-96dpi.png"), 0, 0, height: Pptx.inches(3))
      expect(by_height.width).to eq(Pptx.inches(4))
    end
  end

  it "stretches to fit when both dimensions are given" do
    pic = slide.shapes.add_picture(image("png-96dpi.png"), 0, 0,
                                   width: Pptx.inches(5), height: Pptx.inches(1))
    aggregate_failures do
      expect(pic.width).to eq(Pptx.inches(5))
      expect(pic.height).to eq(Pptx.inches(1))
    end
  end

  it "accepts an IO stream as well as a path" do
    File.open(image("gif.gif"), "rb") do |io|
      pic = slide.shapes.add_picture(io, 0, 0)
      expect(pic.element.xml).to include("descr=\"image.gif\"")
    end
  end

  # The same picture used on several slides should be stored once.
  it "stores one image part however many times the same image is added" do
    other = presentation.slides.add(presentation.slide_layouts[6])
    slide.shapes.add_picture(image("png-96dpi.png"), 0, 0)
    other.shapes.add_picture(image("png-96dpi.png"), 0, 0)
    other.shapes.add_picture(image("gif.gif"), 0, 0)

    media = presentation.part.package.parts.map { |p| p.partname.to_s }
                        .select { |n| n.start_with?("/ppt/media/") }
    expect(media.sort).to eq(["/ppt/media/image1.png", "/ppt/media/image2.gif"])
  end

  it "names the image part after the image's real format" do
    slide.shapes.add_picture(image("tiff-150dpi.tiff"), 0, 0)
    media = presentation.part.package.parts.map { |p| p.partname.to_s }
                        .find { |n| n.start_with?("/ppt/media/") }
    expect(media).to eq("/ppt/media/image1.tiff")
  end
end
