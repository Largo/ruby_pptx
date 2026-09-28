# frozen_string_literal: true

# The rest of the suite runs under whichever backend is active, and CI runs
# it once per backend. These examples cover what only the choice itself does.
RSpec.describe Pptx::Oxml::Backend do
  describe ".select" do
    it "takes the backend named, in any case" do
      aggregate_failures do
        expect(described_class.select("REXML")).to eq(described_class::RexmlBackend)
        expect(described_class.select("nokogiri")).to eq(described_class::NokogiriBackend)
      end
    end

    it "prefers Nokogiri when nothing is named and it can be loaded" do
      expect(described_class.select(nil)).to eq(described_class::NokogiriBackend)
    end

    it "falls back to REXML when Nokogiri cannot be loaded" do
      allow(described_class).to receive(:require).and_call_original
      allow(described_class).to receive(:require).with("nokogiri").and_raise(LoadError)
      expect(described_class.select("")).to eq(described_class::RexmlBackend)
    end

    it "refuses a name it does not know" do
      expect { described_class.select("libxml") }.to raise_error(ArgumentError, /nokogiri or rexml/)
    end
  end

  it "reports the backend in use" do
    expect(%i[nokogiri rexml]).to include(Pptx.xml_backend)
  end

  # Under both backends: a node taken out of its tree still resolves its
  # prefixes, and putting it back does not leave a redundant declaration.
  it "keeps a removed element readable and reinserts it cleanly" do
    root = Pptx::Oxml::Element.parse(
      %(<p:sld #{Pptx::Oxml::Ns.nsdecls("p", "r")}><p:cSld><p:x r:id="rId1"/></p:cSld></p:sld>)
    )
    c_sld = root.find("p:cSld")
    x = c_sld.element_children.first
    c_sld.remove(x)
    aggregate_failures do
      expect(x.nsptag).to eq("p:x")
      expect(x.get("r:id")).to eq("rId1")
      c_sld.append(x)
      expect(root.to_xml).to eq(
        %(<p:sld #{Pptx::Oxml::Ns.nsdecls("p", "r")}><p:cSld><p:x r:id="rId1"/></p:cSld></p:sld>)
      )
    end
  end
end
