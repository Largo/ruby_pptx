# frozen_string_literal: true

RSpec.describe Pptx::Opc::PackURI do
  it "rejects a part name that does not start with a slash" do
    expect { described_class.new("ppt/presentation.xml") }
      .to raise_error(ArgumentError, /must begin with a slash/)
  end

  describe "slicing a part name" do
    subject(:uri) { described_class.new("/ppt/slides/slide21.xml") }

    it { expect(uri.base_uri).to eq("/ppt/slides") }
    it { expect(uri.filename).to eq("slide21.xml") }
    it { expect(uri.ext).to eq("xml") }
    it { expect(uri.idx).to eq(21) }
    it { expect(uri.member_name).to eq("ppt/slides/slide21.xml") }
    it { expect(uri.rels_uri.to_s).to eq("/ppt/slides/_rels/slide21.xml.rels") }
  end

  describe "the package pseudo part name" do
    subject(:uri) { described_class::PACKAGE }

    it { expect(uri.to_s).to eq("/") }
    it { expect(uri.base_uri).to eq("/") }
    it { expect(uri.filename).to eq("") }
    it { expect(uri.idx).to be_nil }
    it { expect(uri.member_name).to eq("") }
    it { expect(uri.rels_uri.to_s).to eq("/_rels/.rels") }
  end

  describe "#idx" do
    it "is nil for a singleton part name" do
      expect(described_class.new("/ppt/presentation.xml").idx).to be_nil
    end

    it "is nil when the stem has no trailing digits" do
      expect(described_class.new("/ppt/tableStyles.xml").idx).to be_nil
    end
  end

  describe ".from_rel_ref" do
    {
      ["/ppt", "slides/slide1.xml"] => "/ppt/slides/slide1.xml",
      ["/ppt/slides", "../media/image1.png"] => "/ppt/media/image1.png",
      ["/ppt/slides", "./slide2.xml"] => "/ppt/slides/slide2.xml",
      ["/", "ppt/presentation.xml"] => "/ppt/presentation.xml",
      ["/ppt/slides", "../../docProps/core.xml"] => "/docProps/core.xml"
    }.each do |(base, ref), expected|
      it "resolves #{ref.inspect} against #{base.inspect}" do
        expect(described_class.from_rel_ref(base, ref).to_s).to eq(expected)
      end
    end

    it "clamps '..' that would escape the package root" do
      expect(described_class.from_rel_ref("/ppt", "../../../etc/passwd").to_s).to eq("/etc/passwd")
    end

    it "does not expand a leading tilde as a home directory" do
      expect(described_class.from_rel_ref("/ppt", "~foo/bar.xml").to_s).to eq("/ppt/~foo/bar.xml")
    end
  end

  describe "#relative_ref" do
    {
      ["/ppt/slideLayouts/slideLayout1.xml", "/ppt/slides"] => "../slideLayouts/slideLayout1.xml",
      ["/ppt/slides/slide1.xml", "/ppt/slides"] => "slide1.xml",
      ["/ppt/presentation.xml", "/"] => "ppt/presentation.xml",
      ["/ppt/media/image1.png", "/ppt/slides"] => "../media/image1.png"
    }.each do |(part, base), expected|
      it "references #{part} from #{base}" do
        expect(described_class.new(part).relative_ref(base)).to eq(expected)
      end
    end
  end

  it "compares and hashes by its string value" do
    aggregate_failures do
      expect(described_class.new("/a.xml")).to eq(described_class.new("/a.xml"))
      expect(described_class.new("/a.xml")).to eq("/a.xml")
      expect(described_class.new("/a.xml")).to be < described_class.new("/b.xml")
      expect({ described_class.new("/a.xml") => 1 }[described_class.new("/a.xml")]).to eq(1)
    end
  end

  it "acts as a String where one is expected" do
    expect(File.join("x", described_class.new("/a.xml"))).to eq("x/a.xml")
    expect("part #{described_class.new("/a.xml")}").to eq("part /a.xml")
  end
end
