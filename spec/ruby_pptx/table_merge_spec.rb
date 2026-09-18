# frozen_string_literal: true

RSpec.describe "merging table cells" do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }

  def table_of(rows, cols, filled: true)
    table = slide.shapes.add_table(rows, cols, at: [0, 0],
                                               size: [Pptx.inches(9), Pptx.inches(3)]).table
    rows.times { |r| cols.times { |c| table[r, c].text = "#{r}#{c}" } } if filled
    table
  end

  subject(:table) { table_of(3, 3) }

  describe "an unmerged table" do
    it "has no origins and nothing spanned" do
      aggregate_failures do
        expect(table.each_cell.map(&:merge_origin?)).to all(be(false))
        expect(table.each_cell.map(&:spanned?)).to all(be(false))
        expect(table[0, 0].span_width).to eq(1)
        expect(table[0, 0].span_height).to eq(1)
      end
    end
  end

  describe "#merge" do
    before { table[0, 0].merge(table[1, 1]) }

    it "makes the top-left cell the origin, carrying the full span" do
      aggregate_failures do
        expect(table[0, 0]).to be_merge_origin
        expect(table[0, 0].span_height).to eq(2)
        expect(table[0, 0].span_width).to eq(2)
      end
    end

    it "marks the covered cells as spanned, not as origins" do
      covered = [[0, 1], [1, 0], [1, 1]].map { |r, c| table[r, c] }
      aggregate_failures do
        expect(covered.map(&:spanned?)).to all(be(true))
        expect(covered.map(&:merge_origin?)).to all(be(false))
      end
    end

    # The merged cell shows one block of text, so the text of everything it
    # covers has to move into the cell that stays visible.
    it "gathers the text into the origin and empties the rest" do
      aggregate_failures do
        expect(table[0, 0].text).to eq("00\n01\n10\n11")
        expect(table[1, 1].text).to eq("")
      end
    end

    it "leaves cells outside the range alone" do
      aggregate_failures do
        expect(table[2, 2].text).to eq("22")
        expect(table[2, 2]).not_to be_spanned
      end
    end

    # Checked through the elements rather than by matching XML text, since
    # attribute order carries no meaning and asserting it would be brittle.
    it "writes the spans and merge flags PowerPoint expects" do
      origin = table[0, 0].element
      aggregate_failures do
        expect(origin.get("rowSpan")).to eq("2")
        expect(origin.get("gridSpan")).to eq("2")
        expect(table[0, 1].element.get("hMerge")).to eq("1")
        expect(table[1, 0].element.get("vMerge")).to eq("1")
        expect(table[1, 1].element.get("hMerge")).to eq("1")
        expect(table[1, 1].element.get("vMerge")).to eq("1")
      end
    end
  end

  describe "merge geometry" do
    # The two cells are opposite corners, so any diagonal in any order names
    # the same rectangle.
    [[[0, 0], [1, 1]], [[1, 1], [0, 0]], [[0, 1], [1, 0]], [[1, 0], [0, 1]]].each do |a, b|
      it "merges the same range given corners #{a.inspect} and #{b.inspect}" do
        table[*a].merge(table[*b])
        aggregate_failures do
          expect(table[0, 0]).to be_merge_origin
          expect(table[0, 0].span_height).to eq(2)
          expect(table[0, 0].span_width).to eq(2)
        end
      end
    end

    it "merges a whole row" do
      table[0, 0].merge(table[0, 2])
      aggregate_failures do
        expect(table[0, 0].span_width).to eq(3)
        expect(table[0, 0].span_height).to eq(1)
      end
    end

    it "merges a whole column" do
      table[0, 0].merge(table[2, 0])
      aggregate_failures do
        expect(table[0, 0].span_height).to eq(3)
        expect(table[0, 0].span_width).to eq(1)
      end
    end
  end

  describe "refusing an impossible merge" do
    it "will not merge over an existing merge" do
      table[0, 0].merge(table[1, 1])
      expect { table[0, 0].merge(table[2, 2]) }
        .to raise_error(Pptx::Error, /already contains a merged cell/)
    end

    it "will not merge across tables" do
      other = table_of(2, 2)
      expect { table[0, 0].merge(other[1, 1]) }
        .to raise_error(Pptx::Error, /different tables/)
    end
  end

  describe "#split" do
    it "gives back a separate cell for every position the merge covered" do
      table[0, 0].merge(table[1, 1])
      table[0, 0].split
      aggregate_failures do
        expect(table[0, 0]).not_to be_merge_origin
        expect(table[0, 0].span_height).to eq(1)
        expect(table[1, 1]).not_to be_spanned
        expect(table.each_cell.map(&:spanned?)).to all(be(false))
      end
    end

    # Splitting cannot put the text back where it came from, because the merge
    # discarded which cell each paragraph belonged to.
    it "leaves the gathered text in the origin cell" do
      table[0, 0].merge(table[1, 1])
      table[0, 0].split
      aggregate_failures do
        expect(table[0, 0].text).to eq("00\n01\n10\n11")
        expect(table[1, 1].text).to eq("")
      end
    end

    it "refuses on a cell that is not a merge origin" do
      expect { table[2, 2].split }
        .to raise_error(Pptx::Error, /only a merge-origin cell can be split/)
    end

    it "refuses on a spanned cell" do
      table[0, 0].merge(table[1, 1])
      expect { table[1, 1].split }.to raise_error(Pptx::Error, /merge-origin/)
    end
  end

  describe "text gathering" do
    it "does not leave a leading blank line when the origin starts empty" do
      blank = table_of(2, 2, filled: false)
      blank[0, 1].text = "value"
      blank[0, 0].merge(blank[1, 1])
      expect(blank[0, 0].text).to eq("value")
    end

    it "contributes nothing from empty cells" do
      blank = table_of(2, 2, filled: false)
      blank[0, 0].text = "only"
      blank[0, 0].merge(blank[1, 1])
      expect(blank[0, 0].text).to eq("only")
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "merges and splits identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          t = slide.shapes.add_table(3, 3, 0, 0, Inches(9), Inches(3)).table
          for r in range(3):
              for c in range(3):
                  t.cell(r, c).text = "%d%d" % (r, c)
          t.cell(0, 0).merge(t.cell(1, 1))
          t.cell(2, 0).merge(t.cell(2, 2))
          t.cell(2, 0).split()
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          t = slide.shapes.add_table(3, 3, at: [0, 0],
                                           size: [Pptx.inches(9), Pptx.inches(3)]).table
          3.times { |r| 3.times { |c| t[r, c].text = "#{r}#{c}" } }
          t[0, 0].merge(t[1, 1])
          t[2, 0].merge(t[2, 2])
          t[2, 0].split
          prs.save(path)
        }
      )
    end
  end
end
