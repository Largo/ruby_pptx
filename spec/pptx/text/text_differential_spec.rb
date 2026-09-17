# frozen_string_literal: true

RSpec.describe "text and DrawingML agreement with python-pptx" do
  before { require_oracle! }

  # The M5 exit criterion. Text is where the most XML is generated per call,
  # and where run-splitting, escaping and formatting all have to line up.
  it "writes a formatted deck identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Pt, Inches
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        slide.shapes.title.text = "Quarterly Review"

        body = slide.placeholders[1].text_frame
        body.text = "Revenue up 12%\\nCosts flat\\vsecond line"

        p0 = body.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        p0.level = 0
        p0.line_spacing = 1.5
        p0.space_after = Pt(12)
        run = p0.runs[0]
        run.font.bold = True
        run.font.italic = False
        run.font.size = Pt(24)
        run.font.name = "Courier New"
        run.font.underline = True
        run.font.color.rgb = RGBColor(0xC0, 0x50, 0x4D)

        p1 = body.paragraphs[1]
        p1.level = 1
        p1.space_before = Pt(6)
        p1.runs[0].font.color.theme_color = 5

        body.word_wrap = True
        body.vertical_anchor = MSO_ANCHOR.MIDDLE
        body.margin_left = Inches(0.25)

        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[1])

        slide.shapes.title.text = "Quarterly Review"

        body = slide.placeholders[1].text_frame
        body.text = "Revenue up 12%\nCosts flat\vsecond line"

        p0 = body.paragraphs[0]
        p0.alignment = Pptx::Enum::PP_ALIGN::CENTER
        p0.level = 0
        p0.line_spacing = 1.5
        p0.space_after = Pptx.pt(12)
        run = p0.runs[0]
        run.font.bold = true
        run.font.italic = false
        run.font.size = Pptx.pt(24)
        run.font.name = "Courier New"
        run.font.underline = true
        run.font.color.rgb = Pptx::RGBColor.new(0xC0, 0x50, 0x4D)

        p1 = body.paragraphs[1]
        p1.level = 1
        p1.space_before = Pptx.pt(6)
        p1.runs[0].font.color.theme_color = 5

        body.word_wrap = true
        body.vertical_anchor = Pptx::Enum::MSO_ANCHOR::MIDDLE
        body.margin_left = Pptx.inches(0.25)

        prs.save(path)
      }
    )
  end

  it "fills and outlines an auto shape identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Pt, Inches
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE

        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(1))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        shape.line.color.rgb = RGBColor(0xFF, 0xFF, 0x00)
        shape.line.width = Pt(3)
        shape.text_frame.text = "Filled"
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
          Pptx::Enum::MSO_SHAPE::ROUNDED_RECTANGLE,
          Pptx.inches(1), Pptx.inches(1), Pptx.inches(3), Pptx.inches(1)
        )
        shape.fill.solid
        shape.fill.fore_color.rgb = Pptx::RGBColor.new(0x1F, 0x49, 0x7D)
        shape.line.color.rgb = Pptx::RGBColor.new(0xFF, 0xFF, 0x00)
        shape.line.width = Pptx.pt(3)
        shape.text_frame.text = "Filled"
        prs.save(path)
      }
    )
  end

  it "makes a shape background-filled identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Inches
        from pptx.enum.shapes import MSO_SHAPE
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(1), Inches(1), Inches(2), Inches(2))
        shape.fill.background()
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(
          Pptx::Enum::MSO_SHAPE::OVAL,
          Pptx.inches(1), Pptx.inches(1), Pptx.inches(2), Pptx.inches(2)
        )
        shape.fill.background
        prs.save(path)
      }
    )
  end

  it "writes a text box with several runs identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Inches, Pt
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tf = tb.text_frame
        tf.text = "Plain "
        r = tf.paragraphs[0].add_run()
        r.text = "bold"
        r.font.bold = True
        p2 = tf.add_paragraph()
        p2.text = "second paragraph"
        p2.font.size = Pt(14)
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        tb = slide.shapes.add_textbox(Pptx.inches(1), Pptx.inches(1),
                                      Pptx.inches(4), Pptx.inches(1))
        tf = tb.text_frame
        tf.text = "Plain "
        r = tf.paragraphs[0].add_run
        r.text = "bold"
        r.font.bold = true
        p2 = tf.add_paragraph
        p2.text = "second paragraph"
        p2.font.size = Pptx.pt(14)
        prs.save(path)
      }
    )
  end
end
