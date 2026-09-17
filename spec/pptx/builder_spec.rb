# frozen_string_literal: true

RSpec.describe Pptx::DeckBuilder do
  describe "Pptx.build" do
    it "returns a presentation based on the default template" do
      deck = Pptx.build
      aggregate_failures do
        expect(deck).to be_a(Pptx::Presentation)
        expect(deck.slides).to be_empty
      end
    end

    it "starts from a template when one is given" do
      Tempfile.create(["template", ".pptx"]) do |file|
        file.close
        Pptx::Presentation.new_default.tap { |p| p.slides.add(p.slide_layouts[6]) }.save(file.path)
        expect(Pptx.build(file.path).slides.size).to eq(1)
      end
    end

    it "sets the slide size by name or by explicit dimensions" do
      aggregate_failures do
        expect(Pptx.build { |d| d.slide_size = :widescreen }.slide_width.inches)
          .to be_within(1e-3).of(13.333)
        expect(Pptx.build { |d| d.slide_size = :standard }.slide_width).to eq(Pptx.inches(10))
        expect(Pptx.build { |d| d.slide_size = [Pptx.inches(12), Pptx.inches(9)] }.slide_height)
          .to eq(Pptx.inches(9))
      end
    end
  end

  describe "#slide" do
    it "defaults to the Title and Content layout" do
      deck = Pptx.build { |d| d.slide }
      expect(deck.slides.first.layout.name).to eq("Title and Content")
    end

    it "accepts a layout by name, by index, or as an object" do
      deck = Pptx.build do |d|
        d.slide("Blank")
        d.slide(0)
        d.slide(d.presentation.slide_layouts["Two Content"])
      end
      expect(deck.slides.map { |s| s.layout.name })
        .to eq(["Blank", "Title Slide", "Two Content"])
    end

    it "raises for a layout that does not exist" do
      expect { Pptx.build { |d| d.slide("No Such Layout") } }
        .to raise_error(Pptx::NotFoundError, /no slide layout/)
    end
  end

  describe "#section" do
    it "puts the slides added inside the block into the section" do
      deck = Pptx.build do |d|
        d.slide("Blank")
        d.section("Detail") do
          d.slide("Blank")
          d.slide("Blank")
        end
        d.slide("Blank")
      end

      aggregate_failures do
        expect(deck.sections.map(&:name)).to eq(["Detail"])
        expect(deck.sections["Detail"].slides.size).to eq(2)
        expect(deck.slides.size).to eq(4)
      end
    end

    it "restores the previous section after the block, even on error" do
      deck = Pptx.build do |d|
        d.section("Outer") do
          begin
            d.section("Inner") { raise "boom" }
          rescue RuntimeError
            nil
          end
          d.slide("Blank")
        end
      end
      expect(deck.sections["Outer"].slides.size).to eq(1)
    end
  end

  describe Pptx::SlideBuilder do
    it "sets placeholder text by role and by idx" do
      deck = Pptx.build do |d|
        d.slide("Title Slide") do |s|
          s.title = "Annual Report"
          s.subtitle = "2026"
        end
        d.slide("Title and Content") do |s|
          s.title = "Highlights"
          s[1] = "Growth"
        end
      end

      aggregate_failures do
        expect(deck.slides[0].shapes.title.text).to eq("Annual Report")
        expect(deck.slides[0].placeholders[1].text).to eq("2026")
        expect(deck.slides[1].placeholders[1].text).to eq("Growth")
      end
    end

    it "raises when the layout has no such placeholder" do
      aggregate_failures do
        expect { Pptx.build { |d| d.slide("Blank") { |s| s.title = "x" } } }
          .to raise_error(Pptx::NotFoundError, /no title placeholder/)
        expect { Pptx.build { |d| d.slide("Blank") { |s| s[7] = "x" } } }
          .to raise_error(Pptx::NotFoundError, /placeholder with idx 7/)
      end
    end

    it "adds a styled text box" do
      deck = Pptx.build do |d|
        d.slide("Blank") do |s|
          s.text "A note", at: [Pptx.inches(1), Pptx.inches(1)],
                           font_size: Pptx.pt(18), bold: true, color: "C0504D",
                           align: Pptx::Enum::PP_ALIGN::CENTER
        end
      end

      box = deck.slides.first.shapes.first
      font = box.text_frame.paragraphs.first.runs.first.font
      aggregate_failures do
        expect(box.text_frame.text).to eq("A note")
        expect(box.left).to eq(Pptx.inches(1))
        expect(font.size).to eq(Pptx.pt(18))
        expect(font.bold).to be(true)
        expect(font.color.rgb).to eq(Pptx::RGBColor["C0504D"])
        expect(box.text_frame.paragraphs.first.alignment).to eq(Pptx::Enum::PP_ALIGN::CENTER)
      end
    end

    it "adds a filled and outlined shape with text" do
      deck = Pptx.build do |d|
        d.slide("Blank") do |s|
          s.shape :ROUNDED_RECTANGLE,
                  at: [Pptx.inches(1), Pptx.inches(1)], size: [Pptx.inches(3), Pptx.inches(1)],
                  fill: "1F497D", line: { color: "FFFF00", width: Pptx.pt(3) },
                  text: "Next steps", bold: true
        end
      end

      shape = deck.slides.first.shapes.first
      aggregate_failures do
        expect(shape.auto_shape_type).to eq(Pptx::Enum::MSO_SHAPE::ROUNDED_RECTANGLE)
        expect(shape.fill.fore_color.rgb).to eq(Pptx::RGBColor["1F497D"])
        expect(shape.line.color.rgb).to eq(Pptx::RGBColor["FFFF00"])
        expect(shape.line.width).to eq(Pptx.pt(3))
        expect(shape.text_frame.text).to eq("Next steps")
      end
    end

    it "accepts :none to make a shape transparent" do
      deck = Pptx.build do |d|
        d.slide("Blank") { |s| s.shape :OVAL, at: [0, 0], size: [100, 100], fill: :none }
      end
      expect(deck.slides.first.shapes.first.fill.type)
        .to eq(Pptx::Enum::MSO_FILL_TYPE::BACKGROUND)
    end

    it "adds a table from rows" do
      deck = Pptx.build do |d|
        d.slide("Blank") do |s|
          s.table [%w[Region Q1], ["East", 12]],
                  at: [Pptx.inches(1), Pptx.inches(1)], size: [Pptx.inches(6), Pptx.inches(2)]
        end
      end

      table = deck.slides.first.shapes.first.table
      aggregate_failures do
        expect(table.row_count).to eq(2)
        expect(table.column_count).to eq(2)
        expect(table.cell(1, 1).text).to eq("12")
        expect(table.first_row?).to be(true)
      end
    end

    it "adds a chart from categories and series" do
      deck = Pptx.build do |d|
        d.slide("Blank") do |s|
          s.chart :COLUMN_CLUSTERED, categories: %w[East West],
                                     series: { "Q1" => [1, 2], "Q2" => [3, 4] },
                                     at: [Pptx.inches(1), Pptx.inches(1)],
                                     size: [Pptx.inches(6), Pptx.inches(4)]
        end
      end

      chart = deck.slides.first.shapes.first.chart
      aggregate_failures do
        expect(chart.plot_type).to eq(:BAR)
        expect(chart.categories).to eq(%w[East West])
        expect(chart.series.map(&:name)).to eq(%w[Q1 Q2])
      end
    end

    it "adds a picture at its native size" do
      image = File.expand_path("../fixtures/images/png-96dpi.png", __dir__)
      deck = Pptx.build do |d|
        d.slide("Blank") { |s| s.picture image, at: [Pptx.inches(1), Pptx.inches(1)] }
      end
      expect(deck.slides.first.shapes.first.width.inches).to be_within(1e-9).of(64 / 96.0)
    end
  end

  # The builder is meant to be a convenience over the ordinary API, not a
  # second implementation. This proves it: the same deck built both ways
  # produces the same package, part for part.
  describe "equivalence with the imperative API" do
    it "produces the same package as the calls it stands in for" do
      built = Tempfile.create(["built", ".pptx"]) do |file|
        file.close
        Pptx.build do |d|
          d.slide_size = :widescreen
          d.slide("Title Slide") do |s|
            s.title = "Annual Report"
            s.subtitle = "Prepared in Ruby"
          end
          d.slide("Blank") do |s|
            s.text "A note", at: [Pptx.inches(1), Pptx.inches(0.5)],
                             font_size: Pptx.pt(18), bold: true, color: "C0504D"
            s.shape :CHEVRON, at: [Pptx.inches(1), Pptx.inches(2)],
                              size: [Pptx.inches(3), Pptx.inches(1)],
                              fill: "1F497D", text: "Next"
            s.table [%w[A B], %w[1 2]], at: [Pptx.inches(5), Pptx.inches(2)],
                                        size: [Pptx.inches(4), Pptx.inches(1.5)]
          end
        end.save(file.path)
        ruby_pptx_manifest(file.path)
      end

      manual = Tempfile.create(["manual", ".pptx"]) do |file|
        file.close
        prs = Pptx::Presentation.new_default
        prs.slide_width = Pptx.inches(13.333)
        prs.slide_height = Pptx.inches(7.5)

        title = prs.slides.add(prs.slide_layouts["Title Slide"])
        title.shapes.title.text = "Annual Report"
        title.placeholders[1].text = "Prepared in Ruby"

        blank = prs.slides.add(prs.slide_layouts["Blank"])
        box = blank.shapes.add_textbox(at: [Pptx.inches(1), Pptx.inches(0.5)], size: [Pptx.inches(4), Pptx.inches(1)])
        box.text_frame.text = "A note"
        box.text_frame.paragraphs.first.runs.first.font.tap do |font|
          font.size = Pptx.pt(18)
          font.bold = true
          font.color.rgb = Pptx::RGBColor["C0504D"]
        end

        shape = blank.shapes.add_shape(Pptx::Enum::MSO_SHAPE::CHEVRON, at: [Pptx.inches(1), Pptx.inches(2)], size: [Pptx.inches(3), Pptx.inches(1)])
        shape.fill.solid
        shape.fill.fore_color.rgb = Pptx::RGBColor["1F497D"]
        shape.text_frame.text = "Next"

        frame = blank.shapes.add_table(2, 2, at: [Pptx.inches(5), Pptx.inches(2)], size: [Pptx.inches(4), Pptx.inches(1.5)])
        table = frame.table
        table.first_row = true
        [%w[A B], %w[1 2]].each_with_index do |row, r|
          row.each_with_index { |value, c| table.cell(r, c).text = value }
        end

        prs.save(file.path)
        ruby_pptx_manifest(file.path)
      end

      expect { compare_manifests(built, manual, []) }.not_to raise_error
    end
  end
end
