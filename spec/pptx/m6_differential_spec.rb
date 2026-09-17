# frozen_string_literal: true

RSpec.describe "pictures and tables agreement with python-pptx" do
  before { skip "python-pptx not importable" unless Pptx::Spec::Differential.oracle_available? }

  def self.images = File.expand_path("../fixtures/images", __dir__)
  def images = self.class.images

  # The M6 exit criterion. A picture adds a binary part, a relationship and a
  # content-type default, so it exercises more of the package machinery than
  # anything before it.
  it "adds a picture identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Inches
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(#{File.join(images, 'png-96dpi.png').inspect},
                                 Inches(1), Inches(1))
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        slide.shapes.add_picture(File.join(images, "png-96dpi.png"),
                                 Pptx.inches(1), Pptx.inches(1))
        prs.save(path)
      }
    )
  end

  it "scales and deduplicates pictures identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Inches
        prs = pptx.Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[6])
        s2 = prs.slides.add_slide(prs.slide_layouts[6])
        png = #{File.join(images, 'png-96dpi.png').inspect}
        gif = #{File.join(images, 'gif.gif').inspect}
        s1.shapes.add_picture(png, Inches(1), Inches(1))
        s1.shapes.add_picture(png, Inches(1), Inches(3), width=Inches(2))
        s2.shapes.add_picture(png, Inches(1), Inches(1), height=Inches(1))
        s2.shapes.add_picture(gif, Inches(4), Inches(1), Inches(2), Inches(2))
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        s1 = prs.slides.add(prs.slide_layouts[6])
        s2 = prs.slides.add(prs.slide_layouts[6])
        png = File.join(images, "png-96dpi.png")
        gif = File.join(images, "gif.gif")
        s1.shapes.add_picture(png, Pptx.inches(1), Pptx.inches(1))
        s1.shapes.add_picture(png, Pptx.inches(1), Pptx.inches(3), width: Pptx.inches(2))
        s2.shapes.add_picture(png, Pptx.inches(1), Pptx.inches(1), height: Pptx.inches(1))
        s2.shapes.add_picture(gif, Pptx.inches(4), Pptx.inches(1),
                              width: Pptx.inches(2), height: Pptx.inches(2))
        prs.save(path)
      }
    )
  end

  it "adds each supported image format identically to python-pptx" do
    names = %w[png-96dpi.png png-nodpi.png jpeg-300dpi.jpg gif.gif bmp.bmp tiff-150dpi.tiff]
    expect_same_package(
      python: <<~PY,
        from pptx.util import Inches
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for i, name in enumerate(#{names.inspect}):
            slide.shapes.add_picture(#{images.inspect} + "/" + name, Inches(1), Inches(i * 0.5))
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        names.each_with_index do |name, i|
          slide.shapes.add_picture(File.join(images, name), Pptx.inches(1), Pptx.inches(i * 0.5))
        end
        prs.save(path)
      }
    )
  end

  it "adds a table identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Inches, Pt
        from pptx.enum.text import MSO_ANCHOR
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        frame = slide.shapes.add_table(3, 4, Inches(0.5), Inches(1.5), Inches(9), Inches(3))
        table = frame.table
        table.cell(0, 0).text = "Region"
        table.cell(0, 1).text = "Q1"
        table.cell(1, 0).text = "EMEA"
        table.cell(1, 1).text = "1,204"
        table.columns[0].width = Inches(3)
        table.rows[0].height = Inches(0.6)
        table.first_col = True
        table.horz_banding = False
        cell = table.cell(0, 0)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.text_frame.paragraphs[0].runs[0].font.bold = True
        cell.text_frame.paragraphs[0].runs[0].font.size = Pt(14)
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        frame = slide.shapes.add_table(3, 4, Pptx.inches(0.5), Pptx.inches(1.5),
                                       Pptx.inches(9), Pptx.inches(3))
        table = frame.table
        table.cell(0, 0).text = "Region"
        table.cell(0, 1).text = "Q1"
        table.cell(1, 0).text = "EMEA"
        table.cell(1, 1).text = "1,204"
        table.columns[0].width = Pptx.inches(3)
        table.rows[0].height = Pptx.inches(0.6)
        table.first_col = true
        table.horz_banding = false
        cell = table.cell(0, 0)
        cell.vertical_anchor = Pptx::Enum::MSO_ANCHOR::MIDDLE
        cell.text_frame.paragraphs[0].runs[0].font.bold = true
        cell.text_frame.paragraphs[0].runs[0].font.size = Pptx.pt(14)
        prs.save(path)
      }
    )
  end

  # A size that does not divide evenly by the row or column count forces the
  # remainder into the last one; an evenly-divisible table never exercises it.
  it "distributes an indivisible table size identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_table(3, 7, 100, 100, 1000001, 700003)
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        slide.shapes.add_table(3, 7, 100, 100, 1_000_001, 700_003)
        prs.save(path)
      }
    )
  end

  it "builds a mixed deck identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Inches, Pt
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.dml.color import RGBColor
        prs = pptx.Presentation()

        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = "Annual Report"
        title_slide.placeholders[1].text_frame.text = "Prepared in Ruby"

        content = prs.slides.add_slide(prs.slide_layouts[1])
        content.shapes.title.text = "Highlights"
        content.placeholders[1].text_frame.text = "Growth\\nMargin\\nOutlook"

        blank = prs.slides.add_slide(prs.slide_layouts[6])
        blank.shapes.add_picture(#{File.join(images, 'jpeg-300dpi.jpg').inspect},
                                 Inches(0.5), Inches(0.5), width=Inches(3))
        box = blank.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(4), Inches(1),
                                     Inches(3), Inches(1))
        box.fill.solid()
        box.fill.fore_color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        box.text_frame.text = "Next"
        t = blank.shapes.add_table(2, 2, Inches(0.5), Inches(4), Inches(6), Inches(2)).table
        t.cell(0, 0).text = "A"
        t.cell(1, 1).text = "B"
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default

        title_slide = prs.slides.add(prs.slide_layouts[0])
        title_slide.shapes.title.text = "Annual Report"
        title_slide.placeholders[1].text_frame.text = "Prepared in Ruby"

        content = prs.slides.add(prs.slide_layouts[1])
        content.shapes.title.text = "Highlights"
        content.placeholders[1].text_frame.text = "Growth\nMargin\nOutlook"

        blank = prs.slides.add(prs.slide_layouts[6])
        blank.shapes.add_picture(File.join(images, "jpeg-300dpi.jpg"),
                                 Pptx.inches(0.5), Pptx.inches(0.5), width: Pptx.inches(3))
        box = blank.shapes.add_shape(Pptx::Enum::MSO_SHAPE::CHEVRON, Pptx.inches(4),
                                     Pptx.inches(1), Pptx.inches(3), Pptx.inches(1))
        box.fill.solid
        box.fill.fore_color.rgb = Pptx::RGBColor.new(0x1F, 0x49, 0x7D)
        box.text_frame.text = "Next"
        t = blank.shapes.add_table(2, 2, Pptx.inches(0.5), Pptx.inches(4),
                                   Pptx.inches(6), Pptx.inches(2)).table
        t.cell(0, 0).text = "A"
        t.cell(1, 1).text = "B"
        prs.save(path)
      }
    )
  end
end
