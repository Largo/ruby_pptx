# frozen_string_literal: true

RSpec.describe Pptx::TextFrame do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts[1]) }
  let(:shape) { slide.shapes.title }
  subject(:frame) { shape.text_frame }

  describe "#text" do
    it "starts empty" do
      expect(frame.text).to eq("")
    end

    it "round-trips a single line" do
      frame.text = "Hello"
      expect(frame.text).to eq("Hello")
    end

    it "makes a paragraph per line feed" do
      frame.text = "One\nTwo\nThree"
      aggregate_failures do
        expect(frame.paragraphs.size).to eq(3)
        expect(frame.paragraphs.map(&:text)).to eq(%w[One Two Three])
      end
    end

    # A vertical tab is how PowerPoint represents a soft carriage return on
    # the clipboard, so it maps to a:br rather than a new paragraph.
    it "makes a line break per vertical tab, within one paragraph" do
      frame.text = "One\vTwo"
      aggregate_failures do
        expect(frame.paragraphs.size).to eq(1)
        expect(frame.paragraphs.first.runs.map(&:text)).to eq(%w[One Two])
        expect(frame.text).to eq("One\vTwo")
      end
    end

    it "replaces rather than appends" do
      frame.text = "First"
      frame.text = "Second"
      aggregate_failures do
        expect(frame.text).to eq("Second")
        expect(frame.paragraphs.size).to eq(1)
      end
    end

    it "keeps one empty paragraph when assigned an empty string" do
      frame.text = ""
      aggregate_failures do
        expect(frame.paragraphs.size).to eq(1)
        expect(frame.text).to eq("")
      end
    end

    it "preserves empty lines" do
      frame.text = "One\n\nThree"
      expect(frame.paragraphs.map(&:text)).to eq(["One", "", "Three"])
    end

    # XML 1.0 cannot carry most control characters, so PowerPoint escapes them
    # as literal text.
    it "escapes control characters other than tab and newline" do
      frame.text = "bell\x07here\tkept"
      aggregate_failures do
        expect(frame.text).to eq("bell_x0007_here\tkept")
        expect(frame.element.xml).to include("_x0007_")
      end
    end

    it "escapes XML markup characters rather than emitting them raw" do
      frame.text = "a < b & c"
      aggregate_failures do
        expect(frame.text).to eq("a < b & c")
        expect(frame.element.xml).to include("&lt;", "&amp;")
      end
    end
  end

  describe "#clear" do
    it "leaves a single empty paragraph" do
      frame.text = "One\nTwo"
      frame.clear
      aggregate_failures do
        expect(frame.paragraphs.size).to eq(1)
        expect(frame.text).to eq("")
      end
    end
  end

  describe "margins and anchoring" do
    it "reports the schema defaults when unset" do
      aggregate_failures do
        expect(frame.margin_left).to eq(Pptx.emu(91_440))
        expect(frame.margin_top).to eq(Pptx.emu(45_720))
        expect(frame.vertical_anchor).to be_nil
      end
    end

    it "round-trips margins and anchor" do
      frame.margin_left = Pptx.inches(0.5)
      frame.vertical_anchor = Pptx::Enum::MSO_ANCHOR::MIDDLE
      aggregate_failures do
        expect(frame.margin_left).to eq(Pptx.inches(0.5))
        expect(frame.vertical_anchor).to eq(Pptx::Enum::MSO_ANCHOR::MIDDLE)
        expect(frame.element.xml).to include('anchor="ctr"')
      end
    end

    it "removes the attribute when a margin is set back to its default" do
      frame.margin_left = Pptx.inches(0.5)
      frame.margin_left = Pptx.emu(91_440)
      expect(frame.element.xml).not_to include("lIns")
    end
  end

  describe "#word_wrap" do
    it "is nil when inherited, and maps true/false to the schema values" do
      aggregate_failures do
        expect(frame.word_wrap).to be_nil
        frame.word_wrap = true
        expect(frame.element.xml).to include('wrap="square"')
        frame.word_wrap = false
        expect(frame.element.xml).to include('wrap="none"')
        frame.word_wrap = nil
        expect(frame.element.xml).not_to include("wrap=")
      end
    end

    it "rejects anything other than true, false or nil" do
      expect { frame.word_wrap = "square" }.to raise_error(ArgumentError, /true, false or nil/)
    end
  end

  describe "#auto_size" do
    it "is nil when unset and round-trips each setting" do
      aggregate_failures do
        expect(frame.auto_size).to be_nil
        frame.auto_size = Pptx::Enum::MSO_AUTO_SIZE::SHAPE_TO_FIT_TEXT
        expect(frame.element.xml).to include("<a:spAutoFit/>")
        frame.auto_size = Pptx::Enum::MSO_AUTO_SIZE::NONE
        expect(frame.element.xml).to include("<a:noAutofit/>")
        expect(frame.element.xml).not_to include("spAutoFit")
        expect(frame.auto_size).to eq(Pptx::Enum::MSO_AUTO_SIZE::NONE)
      end
    end
  end
end

RSpec.describe Pptx::Paragraph do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts[1]) }
  subject(:paragraph) { slide.shapes.title.text_frame.paragraphs.first }

  it "adds runs and reports their text" do
    paragraph.add_run("Hello ")
    paragraph.add_run("world")
    aggregate_failures do
      expect(paragraph.runs.map(&:text)).to eq(["Hello ", "world"])
      expect(paragraph.text).to eq("Hello world")
    end
  end

  # Assigning to a paragraph cannot create another paragraph, so a line feed
  # becomes a break here rather than a split.
  it "turns a line feed into a break rather than a new paragraph" do
    paragraph.text = "One\nTwo"
    aggregate_failures do
      expect(paragraph.runs.map(&:text)).to eq(%w[One Two])
      expect(paragraph.text).to eq("One\vTwo")
    end
  end

  it "round-trips alignment and level" do
    paragraph.alignment = Pptx::Enum::PP_ALIGN::CENTER
    paragraph.level = 2
    aggregate_failures do
      expect(paragraph.alignment).to eq(Pptx::Enum::PP_ALIGN::CENTER)
      expect(paragraph.level).to eq(2)
      expect(paragraph.element.xml).to include('algn="ctr"', 'lvl="2"')
    end
  end

  it "distinguishes line spacing in lines from a fixed distance" do
    aggregate_failures do
      paragraph.line_spacing = 1.5
      expect(paragraph.line_spacing).to eq(1.5)
      expect(paragraph.element.xml).to include("<a:spcPct val=\"150000\"/>")

      paragraph.line_spacing = Pptx.pt(18)
      expect(paragraph.line_spacing).to eq(Pptx.pt(18))
      expect(paragraph.element.xml).to include("<a:spcPts val=\"1800\"/>")
      expect(paragraph.element.xml).not_to include("spcPct")
    end
  end

  it "round-trips space before and after" do
    paragraph.space_before = Pptx.pt(6)
    paragraph.space_after = Pptx.pt(12)
    aggregate_failures do
      expect(paragraph.space_before).to eq(Pptx.pt(6))
      expect(paragraph.space_after).to eq(Pptx.pt(12))
    end
  end

  it "sets one bullet kind at a time, in schema order" do
    paragraph.space_before = Pptx.pt(6)
    paragraph.font.bold = true
    paragraph.bullet = :arabicPeriod
    paragraph.bullet = "–"
    names = paragraph.element.pPr.to_xml.scan(/<a:(\w+)/).flatten - %w[pPr spcPts]
    aggregate_failures do
      expect(paragraph.bullet).to eq("–")
      expect(names).to eq(%w[spcBef buChar defRPr])
    end
  end

  it "suppresses, numbers and inherits bullets" do
    paragraph.bullet = :none
    expect(paragraph.bullet).to eq(:none)
    paragraph.bullet = :arabicPeriod
    expect(paragraph.bullet).to eq(:arabicPeriod)
    paragraph.bullet = nil
    expect(paragraph.bullet).to be_nil
  end

  it "round-trips a hanging indent" do
    paragraph.margin_left = Pptx.pt(18)
    paragraph.indent = -Pptx.pt(18)
    aggregate_failures do
      expect(paragraph.margin_left).to eq(Pptx.pt(18))
      expect(paragraph.indent).to eq(-Pptx.pt(18))
      expect(paragraph.element.pPr.to_xml).to include(%(marL="228600"), %(indent="-228600"))
    end
    paragraph.indent = nil
    expect(paragraph.indent).to be_nil
  end

  it "clears content but keeps paragraph properties" do
    paragraph.alignment = Pptx::Enum::PP_ALIGN::CENTER
    paragraph.text = "gone"
    paragraph.clear
    aggregate_failures do
      expect(paragraph.text).to eq("")
      expect(paragraph.alignment).to eq(Pptx::Enum::PP_ALIGN::CENTER)
    end
  end
end

RSpec.describe Pptx::Font do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts[1]) }
  let(:run) { slide.shapes.title.text_frame.paragraphs.first.add_run("text") }
  subject(:font) { run.font }

  it "reports nil for every inherited property" do
    aggregate_failures do
      expect(font.bold).to be_nil
      expect(font.italic).to be_nil
      expect(font.size).to be_nil
      expect(font.name).to be_nil
      expect(font.underline).to be_nil
    end
  end

  it "round-trips bold, italic and typeface" do
    font.bold = true
    font.italic = false
    font.name = "Courier New"
    aggregate_failures do
      expect(font.bold).to be(true)
      expect(font.italic).to be(false)
      expect(font.name).to eq("Courier New")
      expect(run.element.xml).to include('b="1"', 'i="0"', 'typeface="Courier New"')
    end
  end

  it "restores inheritance when the typeface is set to nil" do
    font.name = "Courier New"
    font.name = nil
    aggregate_failures do
      expect(font.name).to be_nil
      expect(run.element.xml).not_to include("latin")
    end
  end

  it "stores size in centipoints and reads it back as a length" do
    font.size = Pptx.pt(24)
    aggregate_failures do
      expect(font.size).to eq(Pptx.pt(24))
      expect(font.size.pt).to eq(24.0)
      expect(run.element.xml).to include('sz="2400"')
    end
  end

  it "maps underline true and false onto the enumeration" do
    aggregate_failures do
      font.underline = true
      expect(run.element.xml).to include('u="sng"')
      expect(font.underline).to be(true)

      font.underline = false
      expect(run.element.xml).to include('u="none"')
      expect(font.underline).to be(false)

      font.underline = Pptx::Enum::MSO_UNDERLINE::WAVY_LINE
      expect(font.underline).to eq(Pptx::Enum::MSO_UNDERLINE::WAVY_LINE)
    end
  end

  it "makes the fill solid on first access to color" do
    aggregate_failures do
      expect(font.fill.type).to be_nil
      font.color.rgb = Pptx::RGBColor["C0504D"]
      expect(font.fill.type).to eq(Pptx::Enum::MSO_FILL_TYPE::SOLID)
      expect(run.element.xml).to include('<a:srgbClr val="C0504D"/>')
    end
  end

  it "switches a colour from theme to rgb without leaving both" do
    font.color.theme_color = Pptx::Enum::MSO_THEME_COLOR::ACCENT_1
    expect(run.element.xml).to include("schemeClr")
    font.color.rgb = Pptx::RGBColor["FF0000"]
    aggregate_failures do
      expect(run.element.xml).to include("srgbClr")
      expect(run.element.xml).not_to include("schemeClr")
      expect(font.color.theme_color).to be_nil
    end
  end
end

RSpec.describe Pptx::RGBColor do
  it "parses and renders a hex string" do
    aggregate_failures do
      expect(described_class["3C2F80"].to_a).to eq([60, 47, 128])
      expect(described_class.new(60, 47, 128).to_s).to eq("3C2F80")
      expect(described_class["ff0000"].to_s).to eq("FF0000")
    end
  end

  it "rejects malformed input" do
    aggregate_failures do
      expect { described_class["ZZZZZZ"] }.to raise_error(ArgumentError, /hex/)
      expect { described_class["FFF"] }.to raise_error(ArgumentError, /hex/)
      expect { described_class.new(256, 0, 0) }.to raise_error(ArgumentError, /0-255/)
      expect { described_class.new(1.5, 0, 0) }.to raise_error(ArgumentError, /0-255/)
    end
  end

  it "is a value object" do
    aggregate_failures do
      expect(described_class["3C2F80"]).to eq(described_class.new(60, 47, 128))
      expect({ described_class["3C2F80"] => :a }[described_class.new(60, 47, 128)]).to eq(:a)
      expect(described_class["3C2F80"]).to be_frozen
    end
  end
end
