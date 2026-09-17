# frozen_string_literal: true

RSpec.describe Pptx::Presentation do
  def self.fixture = File.expand_path("../fixtures/basic.pptx", __dir__)
  def fixture = self.class.fixture

  subject(:presentation) { described_class.open(fixture) }

  describe "opening" do
    it "opens the built-in default template when given nothing" do
      aggregate_failures do
        expect(described_class.new_default.slides).to be_empty
        expect(described_class.new_default.slide_layouts.size).to eq(11)
      end
    end

    it "opens from an IO stream" do
      File.open(fixture, "rb") do |io|
        expect(described_class.open(io).slides.size).to eq(1)
      end
    end

    it "rejects an OPC package that is not a presentation" do
      # A .docx-shaped package would load as OPC but have the wrong main part.
      allow(Pptx::Package).to receive(:open).and_return(
        instance_double(Pptx::Package, main_document_part:
          instance_double(Pptx::Parts::PresentationPart, content_type: "application/msword"))
      )
      expect { described_class.open("whatever.docx") }
        .to raise_error(Pptx::Error, /not a PowerPoint file/)
    end
  end

  describe "slide size" do
    it "reads width and height as lengths" do
      aggregate_failures do
        expect(presentation.slide_width).to eq(Pptx.inches(10))
        expect(presentation.slide_height).to eq(Pptx.inches(7.5))
      end
    end

    it "accepts a Length or a bare EMU integer" do
      presentation.slide_width = Pptx.inches(13.333)
      presentation.slide_height = 6_858_000
      aggregate_failures do
        expect(presentation.slide_width.inches).to be_within(1e-6).of(13.333)
        expect(presentation.slide_height).to eq(Pptx.emu(6_858_000))
      end
    end

    it "rejects a size outside the range PowerPoint accepts" do
      expect { presentation.slide_width = Pptx.inches(100) }
        .to raise_error(RangeError, /1-56 inches/)
    end
  end

  describe "#slides" do
    it "is Enumerable and knows its size" do
      aggregate_failures do
        expect(presentation.slides.size).to eq(1)
        expect(presentation.slides.map(&:slide_id)).to eq([256])
        expect(presentation.slides.first).to be_a(Pptx::Slide)
      end
    end

    it "indexes from either end and returns nil past the end" do
      aggregate_failures do
        expect(presentation.slides[0]).to eq(presentation.slides.first)
        expect(presentation.slides[-1]).to eq(presentation.slides.first)
        expect(presentation.slides[5]).to be_nil
        expect { presentation.slides.fetch(5) }.to raise_error(IndexError)
      end
    end

    it "finds a slide by its presentation-wide id" do
      aggregate_failures do
        expect(presentation.slides.by_id(256)).to eq(presentation.slides.first)
        expect(presentation.slides.by_id(999)).to be_nil
      end
    end

    it "reports a slide's position" do
      expect(presentation.slides.index(presentation.slides.first)).to eq(0)
    end
  end

  describe "#slide_layouts" do
    it "reads every layout of the first master, in order" do
      expect(presentation.slide_layouts.map(&:name).first(3))
        .to eq(["Title Slide", "Title and Content", "Section Header"])
    end

    it "indexes by position or by name" do
      by_index = presentation.slide_layouts[1]
      aggregate_failures do
        expect(presentation.slide_layouts["Title and Content"]).to eq(by_index)
        expect(presentation.slide_layouts["No Such Layout"]).to be_nil
        expect(presentation.slide_layouts.index(by_index)).to eq(1)
      end
    end

    it "links a layout back to its master" do
      expect(presentation.slide_layouts[0].slide_master).to eq(presentation.slide_master)
    end

    it "reports which slides use a layout" do
      in_use = presentation.slide_layouts["Title and Content"]
      aggregate_failures do
        expect(in_use.used_by_slides.map(&:slide_id)).to eq([256])
        expect(presentation.slide_layouts["Blank"].used_by_slides).to be_empty
      end
    end

    it "refuses to delete a layout that slides still use" do
      expect { presentation.slide_layouts.delete(presentation.slide_layouts["Title and Content"]) }
        .to raise_error(Pptx::Error, /in use by one or more slides/)
    end

    it "deletes an unused layout and drops it from the package" do
      layouts = presentation.slide_layouts
      before = layouts.size
      layouts.delete(layouts["Blank"])
      aggregate_failures do
        expect(layouts.size).to eq(before - 1)
        expect(layouts["Blank"]).to be_nil
      end
    end
  end

  describe "a slide" do
    subject(:slide) { presentation.slides.first }

    it { expect(slide.slide_id).to eq(256) }
    it { expect(slide.layout.name).to eq("Title and Content") }
    it { expect(slide.name).to eq("") }
    it { expect(slide).to be_follows_master_background }
    it { expect(slide).not_to be_has_notes_slide }

    it "round-trips a name through the XML" do
      slide.name = "Opening"
      expect(slide.name).to eq("Opening")
    end
  end

  describe "saving" do
    before { require_oracle! }

    # The M3 exit criterion. Unlike M2 this exercises real re-serialization:
    # the presentation, slide, layout and master parts are now XML parts
    # rebuilt from their element trees rather than blobs passed through.
    it "saves an edited presentation identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          import datetime as dt
          prs = pptx.Presentation(#{fixture.inspect})
          prs.slide_width = 12192000
          prs.slide_height = 6858000
          cp = prs.core_properties
          cp.title = "Ported to Ruby"
          cp.author = "Andi"
          cp.keywords = "ooxml, ruby"
          cp.revision = 7
          cp.modified = dt.datetime(2026, 9, 17, 12, 0, 0)
          cp.created = dt.datetime(2026, 1, 2, 3, 4, 5)
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = described_class.open(fixture)
          prs.slide_width = 12_192_000
          prs.slide_height = 6_858_000
          cp = prs.core_properties
          cp.title = "Ported to Ruby"
          cp.author = "Andi"
          cp.keywords = "ooxml, ruby"
          cp.revision = 7
          cp.modified = Time.utc(2026, 9, 17, 12, 0, 0)
          cp.created = Time.utc(2026, 1, 2, 3, 4, 5)
          prs.save(path)
        }
      )
    end

    it "saves the untouched default template identically to python-pptx" do
      expect_same_package(
        python: "prs = pptx.Presentation(); prs.save(out)",
        ruby: ->(path) { described_class.new_default.save(path) }
      )
    end

    it "reopens what it saved" do
      Tempfile.create(["out", ".pptx"]) do |file|
        file.close
        presentation.tap { |p| p.slide_width = Pptx.inches(13.333) }.save(file.path)
        expect(described_class.open(file.path).slide_width.inches).to be_within(1e-6).of(13.333)
      end
    end
  end
end
