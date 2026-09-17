# frozen_string_literal: true

RSpec.describe Pptx::SlideShapes do
  def self.fixture = File.expand_path("../../fixtures/basic.pptx", __dir__)
  def fixture = self.class.fixture

  let(:presentation) { Pptx::Presentation.new_default }
  let(:layout) { presentation.slide_layouts["Title and Content"] }

  describe "adding a slide" do
    subject(:slide) { presentation.slides.add(layout) }

    it "clones the layout's non-latent placeholders, in z-order" do
      expect(slide.shapes.map(&:name)).to eq(["Title 1", "Content Placeholder 2"])
    end

    it "leaves latent placeholders to the layout" do
      latent = layout.placeholders.map { |ph| ph.element.ph_type.name } &
               %i[DATE FOOTER SLIDE_NUMBER]
      aggregate_failures do
        expect(latent).not_to be_empty
        expect(slide.placeholders.map { |ph| ph.placeholder_format.type.name })
          .to eq(%i[TITLE OBJECT])
      end
    end

    it "gives each new shape a unique id starting after the shape tree's own" do
      expect(slide.shapes.map(&:shape_id)).to eq([2, 3])
    end

    it "carries over idx, type, orientation and size from the layout" do
      cloned = slide.placeholders[1].placeholder_format
      source = layout.placeholders.by_idx(1).placeholder_format
      aggregate_failures do
        expect(cloned.idx).to eq(source.idx)
        expect(cloned.type).to eq(source.type)
        expect(cloned.orientation).to eq(source.orientation)
        expect(cloned.size).to eq(source.size)
      end
    end

    it "gives textual placeholders an empty text body" do
      expect(slide.shapes.title.element.txBody).not_to be_nil
    end

    it "relates the new slide part to its layout" do
      expect(slide.layout).to eq(layout)
    end

    it "registers the slide in the presentation" do
      added = slide
      aggregate_failures do
        expect(presentation.slides.size).to eq(1)
        expect(presentation.slides.first).to eq(added)
        expect(added.slide_id).to eq(256)
      end
    end

    it "numbers a second slide's ids and partname onward" do
      presentation.slides.add(layout)
      second = presentation.slides.add(layout)
      aggregate_failures do
        expect(second.slide_id).to eq(257)
        expect(second.part.partname.to_s).to eq("/ppt/slides/slide2.xml")
      end
    end
  end

  describe "reading shapes from an existing deck" do
    subject(:slide) { Pptx::Presentation.open(fixture).slides.first }

    it "builds a placeholder proxy for each placeholder" do
      expect(slide.shapes.map(&:class)).to all(eq(Pptx::SlidePlaceholder))
    end

    it "finds the title by idx 0, not by position" do
      expect(slide.shapes.title.placeholder_format.idx).to eq(0)
    end

    it "keys placeholders by idx rather than position" do
      aggregate_failures do
        expect(slide.placeholders[0].name).to eq(slide.shapes.title.name)
        expect(slide.placeholders[99]).to be_nil
        expect { slide.placeholders.fetch(99) }.to raise_error(Pptx::NotFoundError)
      end
    end

    it "reads a layout's shapes including the latent placeholders" do
      layout = slide.layout
      expect(layout.placeholders.map { |ph| ph.element.ph_idx }).to eq([0, 1, 10, 11, 12])
    end

    it "reads the master's shapes and placeholders" do
      master = slide.layout.slide_master
      aggregate_failures do
        expect(master.shapes.size).to be > 0
        expect(master.placeholders.map { |ph| ph.class }).to all(eq(Pptx::MasterPlaceholder))
        expect(master.placeholders.by_type(Pptx::Enum::PP_PLACEHOLDER::TITLE)).not_to be_nil
      end
    end
  end

  describe "shape geometry" do
    subject(:shape) { Pptx::Presentation.open(fixture).slides.first.shapes.title }

    it "returns nil for a position the placeholder inherits" do
      expect(shape.left).to be_nil
    end

    it "materializes a transform on first write" do
      shape.left = Pptx.inches(1)
      shape.top = Pptx.inches(2)
      shape.width = Pptx.inches(3)
      shape.height = Pptx.inches(4)
      aggregate_failures do
        expect(shape.left).to eq(Pptx.inches(1))
        expect(shape.top).to eq(Pptx.inches(2))
        expect(shape.width).to eq(Pptx.inches(3))
        expect(shape.height).to eq(Pptx.inches(4))
      end
    end

    it "reports rotation as 0.0 when unset and round-trips a value" do
      aggregate_failures do
        expect(shape.rotation).to eq(0.0)
        shape.rotation = 45.0
        expect(shape.rotation).to eq(45.0)
        expect(shape.element.xfrm.get("rot")).to eq("2700000")
      end
    end

    it "renames the shape through cNvPr" do
      shape.name = "Headline"
      expect(shape.element.xml).to include('name="Headline"')
    end
  end

  describe "placeholder naming" do
    it "numbers a cloned placeholder from its shape id" do
      slide = presentation.slides.add(presentation.slide_layouts["Two Content"])
      expect(slide.shapes.map(&:name))
        .to eq(["Title 1", "Content Placeholder 2", "Content Placeholder 3"])
    end

    it "prefixes a vertical placeholder" do
      slide = presentation.slides.add(presentation.slide_layouts["Vertical Title and Text"])
      expect(slide.shapes.map(&:name)).to eq(["Vertical Title 1", "Vertical Text Placeholder 2"])
    end
  end

  describe "adding shapes" do
    subject(:shapes) { presentation.slides.add(presentation.slide_layouts["Blank"]).shapes }

    it "adds an auto shape named after its type" do
      shape = shapes.add_shape(Pptx::Enum::MSO_SHAPE::ROUNDED_RECTANGLE,
                               Pptx.inches(1), Pptx.inches(1), Pptx.inches(2), Pptx.inches(1))
      aggregate_failures do
        expect(shape.name).to eq("Rounded Rectangle 1")
        expect(shape.shape_type).to eq(Pptx::Enum::MSO_SHAPE_TYPE::AUTO_SHAPE)
        expect(shape.auto_shape_type).to eq(Pptx::Enum::MSO_SHAPE::ROUNDED_RECTANGLE)
        expect(shape.left).to eq(Pptx.inches(1))
        expect(shape.width).to eq(Pptx.inches(2))
      end
    end

    it "accepts a shape type by symbol or MS API value" do
      aggregate_failures do
        expect(shapes.add_shape(:OVAL, 0, 0, 100, 100).auto_shape_type)
          .to eq(Pptx::Enum::MSO_SHAPE::OVAL)
        expect(shapes.add_shape(9, 0, 0, 100, 100).auto_shape_type)
          .to eq(Pptx::Enum::MSO_SHAPE::OVAL)
      end
    end

    it "adds a text box, which is not an auto shape" do
      shape = shapes.add_textbox(Pptx.inches(1), Pptx.inches(3), Pptx.inches(4), Pptx.inches(1))
      aggregate_failures do
        expect(shape.name).to eq("TextBox 1")
        expect(shape.shape_type).to eq(Pptx::Enum::MSO_SHAPE_TYPE::TEXT_BOX)
        expect(shape.auto_shape_type).to be_nil
      end
    end

    it "numbers each added shape from its own id" do
      3.times { shapes.add_textbox(0, 0, 100, 100) }
      expect(shapes.map(&:name)).to eq(["TextBox 1", "TextBox 2", "TextBox 3"])
    end

    it "rejects a shape type that does not exist" do
      expect { shapes.add_shape(:NOT_A_SHAPE, 0, 0, 100, 100) }
        .to raise_error(ArgumentError, /not a member/)
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    # The M4 exit criterion: adding a slide is the first operation that both
    # creates a new part and writes non-trivial shape XML into it.
    it "adds a slide identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          prs = pptx.Presentation()
          prs.slides.add_slide(prs.slide_layouts[1])
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          prs.slides.add(prs.slide_layouts[1])
          prs.save(path)
        }
      )
    end

    it "adds several slides from different layouts identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          prs = pptx.Presentation()
          for idx in (0, 1, 3, 8, 10):
              prs.slides.add_slide(prs.slide_layouts[idx])
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          [0, 1, 3, 8, 10].each { |idx| prs.slides.add(prs.slide_layouts[idx]) }
          prs.save(path)
        }
      )
    end

    it "adds auto shapes and text boxes identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.enum.shapes import MSO_SHAPE
          from pptx.util import Inches
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
          slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(1), Inches(2), Inches(2), Inches(1))
          slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          slide.shapes.add_shape(Pptx::Enum::MSO_SHAPE::ROUNDED_RECTANGLE,
                                 Pptx.inches(1), Pptx.inches(1), Pptx.inches(2), Pptx.inches(1))
          slide.shapes.add_shape(Pptx::Enum::MSO_SHAPE::CHEVRON,
                                 Pptx.inches(1), Pptx.inches(2), Pptx.inches(2), Pptx.inches(1))
          slide.shapes.add_textbox(Pptx.inches(1), Pptx.inches(3),
                                   Pptx.inches(4), Pptx.inches(1))
          prs.save(path)
        }
      )
    end

    it "adds a slide to an existing deck identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          prs = pptx.Presentation(#{fixture.inspect})
          prs.slides.add_slide(prs.slide_layouts[6])
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.open(fixture)
          prs.slides.add(prs.slide_layouts[6])
          prs.save(path)
        }
      )
    end
  end
end
