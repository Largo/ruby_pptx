# frozen_string_literal: true

RSpec.describe "line dash, shadow, adjustments and gradients" do
  let(:deck) { Pptx::Presentation.new_default }
  let(:slide) { deck.slides.add(deck.slide_layouts[6]) }
  let(:shape) do
    slide.shapes.add_shape(:rounded_rectangle, at: [0, 0], size: [Pptx.inches(3), Pptx.inches(1)])
  end

  describe Pptx::LineFormat do
    # Regression: reading the width used to add an `a:ln` to the shape, so a
    # file could change merely by being inspected.
    it "does not write anything when read" do
      shape.line.width
      shape.line.dash_style
      expect(shape.element.spPr.ln).to be_nil
    end

    it "sets and clears a dash style" do
      shape.line.dash_style = :dash_dot
      aggregate_failures do
        expect(shape.line.dash_style.name).to eq(:DASH_DOT)
        shape.line.dash_style = nil
        expect(shape.line.dash_style).to be_nil
      end
    end

    # A custom dash overrides the inherited style just as a preset one does,
    # so restoring inheritance has to remove both.
    it "clears a custom dash as well as a preset one" do
      shape.line.width = Pptx.pt(1)
      shape.element.spPr.ln.get_or_add_custDash
      shape.line.dash_style = nil
      expect(shape.element.spPr.ln.custDash).to be_nil
    end

    it "clears a dash on a shape that has no line at all without adding one" do
      shape.line.dash_style = nil
      expect(shape.element.spPr.ln).to be_nil
    end

    it "refuses a dash style that does not exist" do
      expect { shape.line.dash_style = :wiggly }.to raise_error(ArgumentError)
    end
  end

  describe Pptx::ShadowFormat do
    it "inherits until told otherwise" do
      aggregate_failures do
        expect(shape.shadow).to be_inherit
        shape.shadow.inherit = false
        expect(shape.shadow).not_to be_inherit
        shape.shadow.inherit = true
        expect(shape.shadow).to be_inherit
      end
    end

    it "works on groups, whose effects live in grpSpPr" do
      group = slide.shapes.add_group_shape([shape])
      group.shadow.inherit = false
      expect(group.element.grpSpPr.effectLst).not_to be_nil
    end

    it "declines on a graphic frame, whose shadow belongs to its contents" do
      frame = slide.shapes.add_table(1, 1, at: [0, 0], size: [100, 100])
      expect { frame.shadow }.to raise_error(Pptx::Error, /graphic frame/)
    end
  end

  describe Pptx::Adjustments do
    it "starts at the shape type's defaults" do
      expect(shape.adjustments.to_a).to eq([0.16667])
    end

    it "offers every handle the shape type defines" do
      arc = slide.shapes.add_shape(:block_arc, at: [0, 0], size: [100, 100])
      expect(arc.adjustments.size).to eq(3)
    end

    it "sets a handle and writes it to the geometry" do
      shape.adjustments[0] = 0.4
      aggregate_failures do
        expect(shape.adjustments[0]).to eq(0.4)
        expect(shape.element.spPr.prstGeom.gd_list.map(&:fmla)).to eq(["val 40000"])
      end
    end

    # Truncated, not rounded: 0.29 * 100_000 is 28999.999... in floating
    # point, and python-pptx stores 28999. Rounding would store 29000.
    it "truncates rather than rounds, as python-pptx does" do
      shape.adjustments[0] = 0.29
      expect(shape.element.spPr.prstGeom.gd_list.map(&:fmla)).to eq(["val 28999"])
    end

    # Reading has to come from the file, not from the collection that was
    # just written to: a fresh collection over the same geometry, and a
    # shape reloaded from disk.
    it "reads handle values back from the geometry" do
      shape.adjustments[0] = 0.4
      aggregate_failures do
        expect(Pptx::Adjustments.new(shape.element.spPr.prstGeom)[0]).to eq(0.4)
        reloaded = Tempfile.create(["adj", ".pptx"]) do |file|
          file.close
          deck.save(file.path)
          Pptx::Presentation.open(file.path).slides[0].shapes[0]
        end
        expect(reloaded.adjustments.to_a).to eq([0.4])
      end
    end

    it "ignores a guide that names no handle of the shape type" do
      shape.element.spPr.prstGeom.rewrite_guides([["adj", 40_000], ["bogus", 5]])
      expect(Pptx::Adjustments.new(shape.element.spPr.prstGeom).to_a).to eq([0.4])
    end

    it "writes every handle when one is set" do
      arc = slide.shapes.add_shape(:block_arc, at: [0, 0], size: [100, 100])
      arc.adjustments[1] = 0.5
      expect(arc.element.spPr.prstGeom.gd_list.size).to eq(3)
    end

    it "is empty for a shape with no preset geometry" do
      expect(Pptx::Adjustments.new(nil).to_a).to eq([])
    end

    it "refuses an index past the last handle" do
      expect { shape.adjustments[3] = 0.5 }.to raise_error(IndexError, /1 adjustment, not 4/)
    end

    it "refuses a value that is not a number" do
      expect { shape.adjustments[0] = "0.5" }.to raise_error(ArgumentError, /numeric/)
    end
  end

  describe "gradient fills" do
    # Regression: `fill.gradient` wrote an empty `<a:gradFill/>`, which draws
    # nothing. It now starts from PowerPoint's default, as python-pptx does.
    it "starts from PowerPoint's default gradient" do
      shape.fill.gradient
      expect(shape.fill.gradient_stops.map(&:position)).to eq([0.0, 1.0])
    end

    # python-pptx raises TypeError here: its default `a:lin` has no angle and
    # it subtracts nil from 360. No angle means the direction is inherited,
    # which is what nil says everywhere else in this API.
    it "reports an inherited angle on a fresh gradient rather than raising" do
      shape.fill.gradient
      expect(shape.fill.gradient_angle).to be_nil
    end

    it "gives table cells and backgrounds the same default" do
      cell = slide.shapes.add_table(1, 1, at: [0, 0], size: [100, 100]).table.cell(0, 0)
      cell.fill.gradient
      slide.background.fill.gradient
      expect([cell.fill.gradient_stops.size, slide.background.fill.gradient_stops.size]).to eq([2, 2])
    end

    it "reports the angle counter-clockwise though the file stores it clockwise" do
      shape.fill.gradient
      shape.fill.gradient_angle = 30
      aggregate_failures do
        expect(shape.fill.gradient_angle).to eq(30.0)
        expect(shape.fill.send(:fill_element).lin.ang).to eq(330.0)
      end
    end

    # Stored clockwise as 360 - angle, so 0 must not come back as 360.
    it "reads a zero angle back as zero" do
      shape.fill.gradient
      shape.fill.gradient_angle = 0
      expect(shape.fill.gradient_angle).to eq(0.0)
    end

    it "moves and recolours stops" do
      shape.fill.gradient
      stop = shape.fill.gradient_stops[1]
      stop.position = 0.25
      stop.color.rgb = Pptx::RGBColor["C0504D"]
      aggregate_failures do
        expect(shape.fill.gradient_stops[1].position).to eq(0.25)
        expect(shape.fill.gradient_stops[1].color.rgb.to_s).to eq("C0504D")
      end
    end

    it "refuses gradient properties on a fill that is not a gradient" do
      shape.fill.solid
      aggregate_failures do
        expect { shape.fill.gradient_angle }.to raise_error(Pptx::Error, /not a gradient/)
        expect { shape.fill.gradient_stops }.to raise_error(Pptx::Error, /not a gradient/)
      end
    end

    it "refuses an angle on a radial gradient" do
      shape.fill.gradient
      grad = shape.fill.send(:fill_element)
      grad.remove_lin
      grad.get_or_add_path
      expect { shape.fill.gradient_angle }.to raise_error(Pptx::Error, /not a linear/)
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "writes dash styles and shadows identically" do
      expect_same_package(
        python: <<~PY,
          from pptx.enum.shapes import MSO_SHAPE
          from pptx.enum.dml import MSO_LINE_DASH_STYLE
          from pptx.util import Inches
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          a = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(1), Inches(1))
          b = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2), 0, Inches(1), Inches(1))
          c = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(4), 0, Inches(1), Inches(1))
          a.line.dash_style = MSO_LINE_DASH_STYLE.LONG_DASH_DOT
          b.line.dash_style = MSO_LINE_DASH_STYLE.DASH
          b.line.dash_style = None
          _ = c.line.width
          a.shadow.inherit = False
          b.shadow.inherit = False
          b.shadow.inherit = True
          group = slide.shapes.add_group_shape([c])
          group.shadow.inherit = False
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          a = slide.shapes.add_shape(:rectangle, at: [0, 0], size: [Pptx.inches(1), Pptx.inches(1)])
          b = slide.shapes.add_shape(:oval, at: [Pptx.inches(2), 0], size: [Pptx.inches(1), Pptx.inches(1)])
          c = slide.shapes.add_shape(:oval, at: [Pptx.inches(4), 0], size: [Pptx.inches(1), Pptx.inches(1)])
          a.line.dash_style = :long_dash_dot
          b.line.dash_style = :dash
          b.line.dash_style = nil
          c.line.width
          a.shadow.inherit = false
          b.shadow.inherit = false
          b.shadow.inherit = true
          group = slide.shapes.add_group_shape([c])
          group.shadow.inherit = false
          prs.save(path)
        }
      )
    end

    it "writes adjustments identically, truncation included" do
      expect_same_package(
        python: <<~PY,
          from pptx.enum.shapes import MSO_SHAPE
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          r = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, 0, 0, 914400, 457200)
          r.adjustments[0] = 0.29
          arc = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, 0, 914400, 914400, 914400)
          arc.adjustments[1] = -0.125
          arc.adjustments[2] = 1.75
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          r = slide.shapes.add_shape(:rounded_rectangle, at: [0, 0], size: [914_400, 457_200])
          r.adjustments[0] = 0.29
          arc = slide.shapes.add_shape(:block_arc, at: [0, 914_400], size: [914_400, 914_400])
          arc.adjustments[1] = -0.125
          arc.adjustments[2] = 1.75
          prs.save(path)
        }
      )
    end

    it "reads adjustments and gradient angles python-pptx wrote" do
      result = Tempfile.create(["py-fmt", ".pptx"]) do |file|
        file.close
        script = <<~PY
          import sys, pptx
          from pptx.enum.shapes import MSO_SHAPE
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          arc = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, 0, 0, 914400, 914400)
          arc.adjustments[0] = 0.125
          arc.adjustments[2] = 0.29
          arc.fill.gradient()
          arc.fill.gradient_angle = 30
          prs.save(sys.argv[1])
        PY
        _, err, status = Open3.capture3("python3", "-c", script, file.path)
        raise "python-pptx failed: #{err}" unless status.success?

        arc = Pptx::Presentation.open(file.path).slides[0].shapes[0]
        [arc.adjustments.to_a, arc.fill.gradient_angle]
      end
      expect(result).to eq([[0.125, 0.0, 0.28999], 30.0])
    end

    it "writes gradients identically on shapes, cells and backgrounds" do
      expect_same_package(
        python: <<~PY,
          from pptx.enum.shapes import MSO_SHAPE
          from pptx.dml.color import RGBColor
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, 914400, 914400)
          s.fill.gradient()
          s.fill.gradient_angle = 30
          stop = s.fill.gradient_stops[1]
          stop.position = 0.25
          stop.color.rgb = RGBColor(0xC0, 0x50, 0x4D)
          cell = slide.shapes.add_table(1, 1, 0, 914400, 914400, 914400).table.cell(0, 0)
          cell.fill.gradient()
          slide.background.fill.gradient()
          slide.background.fill.gradient_angle = 0
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          s = slide.shapes.add_shape(:rectangle, at: [0, 0], size: [914_400, 914_400])
          s.fill.gradient
          s.fill.gradient_angle = 30
          stop = s.fill.gradient_stops[1]
          stop.position = 0.25
          stop.color.rgb = Pptx::RGBColor["C0504D"]
          cell = slide.shapes.add_table(1, 1, at: [0, 914_400], size: [914_400, 914_400]).table.cell(0, 0)
          cell.fill.gradient
          slide.background.fill.gradient
          slide.background.fill.gradient_angle = 0
          prs.save(path)
        }
      )
    end
  end
end
