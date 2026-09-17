# frozen_string_literal: true

RSpec.describe Pptx::FreeformBuilder do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }
  # One local unit per hundredth of an inch, so a 100-unit path is an inch.
  let(:scale) { Pptx.inches(1).emu / 100.0 }

  def triangle(**options)
    slide.shapes.add_freeform(at: [Pptx.inches(1), Pptx.inches(1)], scale: scale, **options) do |f|
      f.line_to(100, 0)
      f.line_to(50, 100)
    end
  end

  describe "adding one" do
    subject(:shape) { triangle }

    it "is a freeform shape with custom geometry" do
      aggregate_failures do
        expect(shape.shape_type).to eq(Pptx::Enum::MSO_SHAPE_TYPE::FREEFORM)
        expect(shape.element).to be_custom_geometry
        expect(shape.auto_shape_type).to be_nil
      end
    end

    # The shape's extents come from the path, scaled; nothing is passed in.
    it "sizes itself from the path" do
      aggregate_failures do
        expect(shape.left).to eq(Pptx.inches(1))
        expect(shape.top).to eq(Pptx.inches(1))
        expect(shape.width).to eq(Pptx.inches(1))
        expect(shape.height).to eq(Pptx.inches(1))
      end
    end

    it "writes the pen movements as a path" do
      xml = shape.element.xpath("./p:spPr/a:custGeom/a:pathLst").first.xml.gsub(/\s+/, " ")
      aggregate_failures do
        expect(xml).to include('<a:path w="100" h="100">')
        expect(xml).to include('<a:moveTo> <a:pt x="0" y="0"/> </a:moveTo>')
        expect(xml).to include('<a:lnTo> <a:pt x="100" y="0"/> </a:lnTo>')
        expect(xml).to include("<a:close/>")
      end
    end

    it "leaves the contour open when asked" do
      open = slide.shapes.add_freeform(at: [0, 0], scale: scale, close: false) do |f|
        f.line_to(100, 0)
      end
      expect(open.element.xml).not_to include("<a:close/>")
    end
  end

  describe "coordinate handling" do
    # The path's own units are local; the shape's extents scale them onto the
    # slide, which is what lets the same drawing be placed at any size.
    it "scales local units onto the slide" do
      shape = slide.shapes.add_freeform(at: [0, 0], scale: Pptx.inches(1).emu / 10.0) do |f|
        f.line_to(10, 10)
      end
      aggregate_failures do
        expect(shape.width).to eq(Pptx.inches(1))
        expect(shape.element.xml).to include('<a:path w="10" h="10">')
      end
    end

    it "accepts different horizontal and vertical scales" do
      shape = slide.shapes.add_freeform(at: [0, 0],
                                        scale: [Pptx.inches(1).emu / 10.0,
                                                Pptx.inches(2).emu / 10.0]) do |f|
        f.line_to(10, 10)
      end
      aggregate_failures do
        expect(shape.width).to eq(Pptx.inches(1))
        expect(shape.height).to eq(Pptx.inches(2))
      end
    end

    # The bounding box need not start at the local origin, so points inside
    # the path are measured from the box rather than from the origin.
    it "shifts the path when the drawing does not start at the origin" do
      shape = slide.shapes.add_freeform(at: [Pptx.inches(1), Pptx.inches(1)],
                                        start_x: 50, start_y: 50, scale: scale) do |f|
        f.line_to(150, 50)
        f.line_to(100, 150)
      end
      xml = shape.element.xml.gsub(/\s+/, " ")
      aggregate_failures do
        # The drawing sits half an inch right and down of the placed origin.
        expect(shape.left).to eq(Pptx.inches(1.5))
        expect(shape.top).to eq(Pptx.inches(1.5))
        # ...but the path itself starts at its own corner.
        expect(xml).to include('<a:moveTo> <a:pt x="0" y="0"/> </a:moveTo>')
      end
    end

    it "handles a drawing that goes left and up from its start" do
      shape = slide.shapes.add_freeform(at: [Pptx.inches(2), Pptx.inches(2)],
                                        start_x: 100, start_y: 100, scale: scale) do |f|
        f.line_to(0, 100)
        f.line_to(50, 0)
      end
      xml = shape.element.xml.gsub(/\s+/, " ")
      aggregate_failures do
        expect(shape.left).to eq(Pptx.inches(2))
        expect(shape.width).to eq(Pptx.inches(1))
        expect(xml).to include('<a:moveTo> <a:pt x="100" y="100"/> </a:moveTo>')
      end
    end

    it "rounds fractional coordinates" do
      shape = slide.shapes.add_freeform(at: [0, 0], scale: scale) do |f|
        f.line_to(10.4, 20.6)
      end
      expect(shape.element.xml).to include('<a:pt x="10" y="21"/>')
    end

    # Python rounds half-to-even and Ruby rounds half-up, so a coordinate
    # landing exactly on .5 would otherwise be placed a unit away from where
    # the reference implementation puts it.
    it "rounds a half-way coordinate to even, as python-pptx does" do
      shape = slide.shapes.add_freeform(at: [0, 0], scale: scale) do |f|
        f.line_to(100.5, 99.5)
      end
      expect(shape.element.xml).to include('<a:pt x="100" y="100"/>')
    end
  end

  describe "the builder" do
    subject(:builder) { slide.shapes.build_freeform(scale: scale) }

    it "chains" do
      expect(builder.line_to(10, 0).line_to(0, 10).close).to be(builder)
    end

    it "adds a run of segments at once" do
      builder.add_line_segments([[50, 0], [50, 50], [0, 50]])
      expect(builder.size).to eq(4) # three lines and a close
    end

    it "leaves the run open when asked" do
      builder.add_line_segments([[50, 0]], close: false)
      expect(builder.size).to eq(1)
    end

    # The same geometry can be stamped in several places.
    it "converts to more than one shape" do
      builder.add_line_segments([[50, 0], [50, 50]])
      first = builder.convert_to_shape(origin_at: [Pptx.inches(1), Pptx.inches(1)])
      second = builder.convert_to_shape(origin_at: [Pptx.inches(4), Pptx.inches(1)])
      aggregate_failures do
        expect(first.left).to eq(Pptx.inches(1))
        expect(second.left).to eq(Pptx.inches(4))
        expect(first.width).to eq(second.width)
        expect(slide.shapes.size).to eq(2)
      end
    end

    it "starts a new contour with move_to" do
      builder.line_to(10, 0).close.move_to(20, 0).line_to(30, 0)
      shape = builder.convert_to_shape
      xml = shape.element.xml
      aggregate_failures do
        expect(xml.scan(/<a:moveTo>/).size).to eq(2)
        expect(xml).to include("<a:close/>")
      end
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "builds freeform shapes identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches, Emu
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          scale = Inches(1) / 100

          builder = slide.shapes.build_freeform(0, 0, scale)
          builder.add_line_segments([(100, 0), (50, 100)])
          builder.convert_to_shape(Inches(1), Inches(1))

          # Same geometry, placed again, and one starting away from the origin.
          builder.convert_to_shape(Inches(4), Inches(1))

          offset = slide.shapes.build_freeform(50, 50, scale)
          offset.add_line_segments([(150, 50), (100, 150)], close=False)
          offset.move_to(120, 120)
          offset.add_line_segments([(140.4, 140.6), (100.5, 99.5)])
          offset.convert_to_shape(Inches(1), Inches(3))
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          scale = Pptx.inches(1).emu / 100.0

          builder = slide.shapes.build_freeform(scale: scale)
          builder.add_line_segments([[100, 0], [50, 100]])
          builder.convert_to_shape(origin_at: [Pptx.inches(1), Pptx.inches(1)])
          builder.convert_to_shape(origin_at: [Pptx.inches(4), Pptx.inches(1)])

          offset = slide.shapes.build_freeform(start_x: 50, start_y: 50, scale: scale)
          offset.add_line_segments([[150, 50], [100, 150]], close: false)
          offset.move_to(120, 120)
          offset.add_line_segments([[140.4, 140.6], [100.5, 99.5]])
          offset.convert_to_shape(origin_at: [Pptx.inches(1), Pptx.inches(3)])
          prs.save(path)
        }
      )
    end
  end
end
