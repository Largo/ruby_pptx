# frozen_string_literal: true

require "stringio"

# What a corporate master needs beyond shapes and text styles: placeholder
# defaults on a layout, fields, vertical text, capitals, footer placeholders
# on slides, and the userDrawn mark PowerPoint puts on layout shapes.
RSpec.describe "building a corporate template" do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }
  let(:frame) { slide.shapes.add_textbox(at: inches(1, 1), size: inches(4, 1)).text_frame }

  def inches(*values) = values.map { |v| Pptx.inches(v) }

  def reopen
    io = StringIO.new
    presentation.save(io)
    Pptx::Presentation.open(StringIO.new(io.string))
  end

  describe "a layout placeholder's list style" do
    let(:layout) { presentation.slide_layouts["Title and Content"] }
    let(:title) { layout.placeholders.find { |ph| ph.element.ph_type == Pptx::Enum::PP_PLACEHOLDER::TITLE } }

    it "sets per-level defaults ahead of the paragraphs, and keeps them across a save" do
      l1 = title.text_frame.list_style.level(0)
      l1.alignment = Pptx::Enum::PP_ALIGN::LEFT
      l1.font.size = Pptx.pt(42.67)
      l1.font.caps = :all
      xml = title.element.find("p:txBody").to_xml
      names = xml.scan(/<a:(bodyPr|lstStyle|p)\b/).flatten
      aggregate_failures do
        expect(names.first(3)).to eq(%w[bodyPr lstStyle p])
        reread = reopen.slide_layouts["Title and Content"].placeholders
                       .find { |ph| ph.element.ph_type == Pptx::Enum::PP_PLACEHOLDER::TITLE }
        style = reread.text_frame.list_style.level(0)
        expect([style.alignment, style.font.caps]).to eq([Pptx::Enum::PP_ALIGN::LEFT, :all])
      end
    end
  end

  describe "fields" do
    it "adds a slide number with its placeholder text before the end-paragraph properties" do
      para = frame.paragraphs.first
      para.font.size = Pptx.pt(16)
      para.element.get_or_add_endParaRPr
      para.add_field(:slide_number)
      xml = para.element.to_xml
      aggregate_failures do
        expect(xml).to match(/<a:fld id="\{[0-9A-F-]{36}\}" type="slidenum"><a:t>/)
        expect(xml.index("a:fld")).to be < xml.index("a:endParaRPr")
        expect(para.text).to eq("‹#›")
      end
    end

    it "adds a date or any other field type" do
      para = frame.paragraphs.first
      para.add_field(:date_time).add_field("datetime4")
      expect(para.element.to_xml.scan(/type="(\w+)"/).flatten).to eq(%w[datetime1 datetime4])
    end

    it "rejects an unknown symbol" do
      expect { frame.paragraphs.first.add_field(:page) }.to raise_error(ArgumentError, /unknown field/)
    end
  end

  describe "text direction" do
    it "runs text bottom to top" do
      frame.text_direction = :vertical_270
      aggregate_failures do
        expect(frame.text_direction).to eq(:vertical_270)
        expect(frame.element.bodyPr.to_xml).to include(%(vert="vert270"))
      end
      frame.text_direction = nil
      expect(frame.text_direction).to be_nil
    end

    it "rejects an unknown direction" do
      expect { frame.text_direction = :diagonal }.to raise_error(ArgumentError, /text_direction/)
    end
  end

  describe "capitals" do
    it "sets all caps, small caps, none and inherit" do
      font = frame.paragraphs.first.add_run("Alos").font
      %i[all small none].each do |value|
        font.caps = value
        expect(font.caps).to eq(value)
      end
      font.caps = nil
      expect(font.caps).to be_nil
    end
  end

  describe "footer placeholders on a slide" do
    let(:layout) { presentation.slide_layouts["Title and Content"] }
    let(:footer_types) { Pptx::SlideLayout::LATENT_PLACEHOLDER_TYPES }

    def footer_placeholders(target)
      target.placeholders.select { |ph| footer_types.include?(ph.element.ph_type) }
    end

    it "leaves them out by default" do
      expect(footer_placeholders(presentation.slides.add(layout))).to be_empty
    end

    it "copies them with their text when asked" do
      numbered = presentation.slides.add(layout, footers: true)
      number = footer_placeholders(numbered)
               .find { |ph| ph.element.ph_type == Pptx::Enum::PP_PLACEHOLDER::SLIDE_NUMBER }
      aggregate_failures do
        expect(footer_placeholders(numbered).size).to eq(footer_placeholders(layout).size)
        expect(number.element.to_xml).to include(%(type="slidenum"))
      end
    end
  end

  describe "userDrawn" do
    it "marks shapes drawn on a layout or master, and not on a slide" do
      layout_box = presentation.slide_layouts["Blank"].shapes
                               .add_shape(:rectangle, at: inches(0, 0), size: inches(1, 1))
      master_line = presentation.slide_masters.first.shapes
                                .add_connector(:straight, begin_at: inches(0, 1), end_at: inches(5, 1))
      slide_box = slide.shapes.add_shape(:rectangle, at: inches(0, 0), size: inches(1, 1))
      aggregate_failures do
        expect(layout_box.element.to_xml).to include(%(<p:nvPr userDrawn="1"/>))
        expect(master_line.element.to_xml).to include(%(<p:nvPr userDrawn="1"/>))
        expect(slide_box.element.to_xml).not_to include("userDrawn")
      end
    end
  end
end
