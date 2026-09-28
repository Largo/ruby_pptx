# frozen_string_literal: true

RSpec.describe Pptx::Oxml::Ns do
  let(:a_uri) { "http://schemas.openxmlformats.org/drawingml/2006/main" }
  let(:p_uri) { "http://schemas.openxmlformats.org/presentationml/2006/main" }

  describe ".qn" do
    it "builds a Clark-notation name" do
      expect(described_class.qn("p:cSld")).to eq("{#{p_uri}}cSld")
    end

    it "raises on an unknown prefix" do
      expect { described_class.qn("zz:foo") }.to raise_error(KeyError, /zz/)
    end

    it "raises on a tag without a prefix" do
      expect { described_class.qn("cSld") }.to raise_error(ArgumentError, /prefixed tag/)
    end
  end

  describe ".prefixed_tag" do
    it "round-trips with .qn" do
      expect(described_class.prefixed_tag(described_class.qn("a:solidFill"))).to eq("a:solidFill")
    end

    it "raises on input that is not Clark notation" do
      expect { described_class.prefixed_tag("p:cSld") }.to raise_error(ArgumentError, /Clark/)
    end
  end

  describe ".namespaces" do
    it "returns the requested subset as a prefix => URI hash" do
      expect(described_class.namespaces("a", "p")).to eq("a" => a_uri, "p" => p_uri)
    end
  end

  describe ".nsdecls" do
    it "renders literal xmlns attribute text" do
      expect(described_class.nsdecls("a")).to eq(%(xmlns:a="#{a_uri}"))
    end
  end

  describe ".clark_name_of" do
    it "derives the registry key from a parsed node" do
      xml = %(<p:cSld #{described_class.nsdecls("p")}/>)
      node = Pptx::Oxml::Element.parse(xml).node
      expect(described_class.clark_name_of(node)).to eq("{#{p_uri}}cSld")
    end

    it "falls back to the bare name for a node with no namespace" do
      expect(described_class.clark_name_of(Pptx::Oxml::Element.parse("<foo/>").node)).to eq("foo")
    end
  end

  it "maps every URI back to a prefix" do
    described_class::NSMAP.each_value do |uri|
      expect(described_class::PFXMAP).to have_key(uri)
    end
  end
end
