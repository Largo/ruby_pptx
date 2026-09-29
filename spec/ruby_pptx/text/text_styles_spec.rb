# frozen_string_literal: true

require "stringio"

# A master's p:txStyles: the per-level formatting every title, body
# placeholder and other text under the master inherits. The values are those
# of a real corporate template.
RSpec.describe Pptx::TextStyles do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:master) { presentation.slide_masters.first }
  let(:styles) { master.text_styles }

  def xml(style) = style.element.to_xml

  it "styles the title level" do
    title = styles.title.level(0)
    title.alignment = Pptx::Enum::PP_ALIGN::CENTER
    title.font.size = Pptx.pt(53)
    title.font.name = "Arial"
    aggregate_failures do
      expect(title.alignment).to eq(Pptx::Enum::PP_ALIGN::CENTER)
      expect(title.font.size).to eq(Pptx.pt(53))
      expect(title.font.name).to eq("Arial")
    end
  end

  it "gives each body level its bullet, hanging indent and size" do
    { 0 => ["•", 0.50, -0.50, 37], 1 => ["–", 1.08, -0.42, 32], 2 => ["•", 1.67, -0.33, 26] }
      .each do |level, (char, mar_l, indent, size)|
        style = styles.body.level(level)
        style.bullet = char
        style.margin_left = Pptx.inches(mar_l)
        style.indent = Pptx.inches(indent)
        style.font.size = Pptx.pt(size)
      end
    aggregate_failures do
      l2 = styles.body.level(1)
      expect([l2.bullet, l2.margin_left, l2.indent, l2.font.size])
        .to eq(["–", Pptx.inches(1.08), Pptx.inches(-0.42), Pptx.pt(32)])
      expect(xml(styles.body.level(2))).to include(%(<a:buChar char="•"/>))
    end
  end

  # The default template already carries a:buFont and a:defRPr on each level;
  # the bullet has to land between them.
  it "keeps the schema order when replacing a bullet in an existing level" do
    styles.body.level(0).bullet = :none
    names = xml(styles.body.level(0)).scan(/<a:(\w+)/).flatten
    expect(names.index("buNone")).to be_between(names.index("buFont"), names.index("defRPr"))
  end

  it "creates missing levels in schema order" do
    other = styles.other
    other.level(4).font.size = Pptx.pt(10)
    other.level(0).font.size = Pptx.pt(18)
    other.default.alignment = Pptx::Enum::PP_ALIGN::LEFT
    names = other.element.to_xml.scan(/<a:(defPPr|lvl\dpPr)/).flatten.uniq
    expect(names.first(2)).to eq(%w[defPPr lvl1pPr])
    expect(names.index("lvl5pPr")).to be > names.index("lvl1pPr")
  end

  it "rejects a level outside 0 to 8" do
    expect { styles.body.level(9) }.to raise_error(ArgumentError, /0 to 8/)
  end

  it "works on a master built in code and survives a save" do
    built = presentation.slide_masters.add(name: "Corporate")
    built.text_styles.body.level(0).bullet = "•"
    built.text_styles.body.level(0).font.size = Pptx.pt(37)
    io = StringIO.new
    presentation.save(io)
    reopened = Pptx::Presentation.open(StringIO.new(io.string)).slide_masters.to_a.last
    l1 = reopened.text_styles.body.level(0)
    expect([l1.bullet, l1.font.size]).to eq(["•", Pptx.pt(37)])
  end
end
