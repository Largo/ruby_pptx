# frozen_string_literal: true

RSpec.describe Pptx::ColorFormat do
  let(:deck) { Pptx::Presentation.new_default }
  let(:shape) do
    deck.slides.add(deck.slide_layouts["Blank"]).shapes
        .add_shape(:rectangle, at: [0, 0], size: [Pptx.cm(2), Pptx.cm(1)])
  end

  describe "#rgb=" do
    # A hex string is the short form: `rgb = "E8722A"` rather than
    # `rgb = Pptx::RGBColor["E8722A"]`.
    it "takes a hex string as well as an RGBColor" do
      shape.fill.solid
      shape.fill.fore_color.rgb = "E8722A"
      expect(shape.fill.fore_color.rgb).to eq(Pptx::RGBColor.new(0xE8, 0x72, 0x2A))
    end

    it "writes the hex uppercase, as OOXML does, whatever case it was given" do
      shape.fill.solid
      shape.fill.fore_color.rgb = "e8722a"
      expect(shape.element.xpath(".//a:srgbClr/@val").map(&:value)).to eq(["E8722A"])
    end

    it "works on a font's colour the same way" do
      shape.text_frame.text = "Chunky"
      font = shape.text_frame.paragraphs[0].runs[0].font
      font.color.rgb = "1F497D"
      expect(font.color.rgb.to_s).to eq("1F497D")
    end

    it "refuses a string that is not six hex digits" do
      shape.fill.solid
      aggregate_failures do
        expect { shape.fill.fore_color.rgb = "E8722" }.to raise_error(ArgumentError, /six-digit hex/)
        expect { shape.fill.fore_color.rgb = "orange" }.to raise_error(ArgumentError, /six-digit hex/)
      end
    end

    it "refuses anything else" do
      shape.fill.solid
      expect { shape.fill.fore_color.rgb = 0xE8722A }.to raise_error(TypeError, /RGBColor/)
    end
  end
end
