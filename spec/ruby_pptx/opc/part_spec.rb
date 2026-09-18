# frozen_string_literal: true

RSpec.describe Pptx::Opc::Part do
  let(:partname) { Pptx::Opc::PackURI.new("/ppt/media/image1.png") }

  subject(:part) { described_class.new(partname, "image/png", nil, "\x89PNG binary".b) }

  it "returns its blob unchanged" do
    expect(part.blob).to eq("\x89PNG binary".b)
  end

  it "returns an empty blob rather than nil when it has none" do
    expect(described_class.new(partname, "image/png", nil).blob).to eq("")
  end

  it "refuses a partname that is not a PackURI" do
    expect { part.partname = "/ppt/media/image2.png" }.to raise_error(TypeError, /PackURI/)
  end

  it "starts with an empty relationship collection based at its own directory" do
    aggregate_failures do
      expect(part.rels).to be_empty
      expect(part.rels.get_or_add("t", part)).to eq("rId1")
      expect(part.rels.xml).to include('Target="image1.png"')
    end
  end
end

RSpec.describe Pptx::Opc::XmlPart do
  let(:partname) { Pptx::Opc::PackURI.new("/ppt/slides/slide1.xml") }
  let(:ns) { Pptx::Oxml::Ns.nsdecls("p", "r") }

  def part_with(inner)
    described_class.load(partname, "application/xml", nil, %(<p:sld #{ns}>#{inner}</p:sld>))
  end

  it "parses its blob into an element tree" do
    expect(part_with("<p:cSld/>").element.nsptag).to eq("p:sld")
  end

  it "reserializes from the element, with an XML declaration" do
    part = part_with("<p:cSld/>")
    aggregate_failures do
      expect(part.blob).to start_with("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>")
      expect(part.blob).to include("<p:cSld/>")
    end
  end

  it "reflects later edits to the element in its blob" do
    part = part_with("<p:cSld/>")
    part.element.append(part.element.build("p:clrMapOvr"))
    expect(part.blob).to include("clrMapOvr")
  end

  it "is its own part, ending the delegation chain" do
    part = part_with("")
    expect(part.part).to equal(part)
  end

  describe "#drop_rel" do
    it "drops a relationship the XML does not reference" do
      part = part_with("<p:cSld/>")
      r_id = part.rels.get_or_add("t", part)
      part.drop_rel(r_id)
      expect(part.rels).to be_empty
    end

    it "drops a relationship referenced exactly once" do
      part = part_with(%(<p:cSld r:id="rId1"/>))
      part.rels.get_or_add("t", part)
      part.drop_rel("rId1")
      expect(part.rels).to be_empty
    end

    # Two references mean something else still depends on it, so removing the
    # relationship would leave dangling r:id attributes.
    it "keeps a relationship referenced more than once" do
      part = part_with(%(<p:cSld r:id="rId1"/><p:clrMapOvr r:id="rId1"/>))
      part.rels.get_or_add("t", part)
      part.drop_rel("rId1")
      expect(part.rels.size).to eq(1)
    end
  end
end

RSpec.describe Pptx::Opc::PartFactory do
  around do |example|
    saved = described_class.registered
    example.run
    described_class.reset!(saved)
  end

  it "builds a plain Part for an unregistered content type" do
    part = described_class.build(
      Pptx::Opc::PackURI.new("/x/y.bin"), "application/unknown", nil, "bytes"
    )
    expect(part).to be_an_instance_of(Pptx::Opc::Part)
  end

  it "builds the registered class for a known content type" do
    klass = Class.new(Pptx::Opc::Part)
    described_class.register("application/custom", klass)
    part = described_class.build(Pptx::Opc::PackURI.new("/x/y.bin"), "application/custom", nil, "b")
    expect(part).to be_an_instance_of(klass)
  end
end

RSpec.describe "relationship loading robustness" do
  let(:parts) do
    { "/ppt/slides/slide1.xml" =>
        Pptx::Opc::Part.new(Pptx::Opc::PackURI.new("/ppt/slides/slide1.xml"), "x", nil) }
  end

  def rels_xml(*relationships)
    Pptx::Oxml::Element.parse(
      "<Relationships xmlns=\"#{Pptx::Opc::NAMESPACE::OPC_RELATIONSHIPS}\">#{relationships.join}</Relationships>"
    )
  end

  # A PowerPoint plugin sometimes "deletes" a relationship by voiding its
  # target instead of removing the element. Refusing to open such a file would
  # be worse than ignoring the dead link.
  it "skips a relationship whose target is not in the package" do
    xml = rels_xml(
      %(<Relationship Id="rId1" Type="t" Target="slides/slide1.xml"/>),
      %(<Relationship Id="rId2" Type="t" Target="slides/NULL"/>)
    )
    rels = Pptx::Opc::Relationships.new("/ppt").load_from_xml("/ppt", xml, parts)
    expect(rels.keys).to eq(["rId1"])
  end

  it "keeps an external relationship even though it has no target part" do
    xml = rels_xml(
      %(<Relationship Id="rId1" Type="t" Target="http://example.com" TargetMode="External"/>)
    )
    rels = Pptx::Opc::Relationships.new("/ppt").load_from_xml("/ppt", xml, parts)
    aggregate_failures do
      expect(rels.keys).to eq(["rId1"])
      expect(rels.fetch("rId1")).to be_external
      expect { rels.fetch("rId1").target_part }.to raise_error(Pptx::Error, /external/)
    end
  end

  it "replaces any relationships already loaded" do
    rels = Pptx::Opc::Relationships.new("/ppt")
    rels.get_or_add("t", parts.values.first)
    rels.load_from_xml("/ppt", rels_xml, parts)
    expect(rels).to be_empty
  end
end
