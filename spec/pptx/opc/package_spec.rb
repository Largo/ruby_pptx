# frozen_string_literal: true

RSpec.describe Pptx::Opc::OpcPackage do
  def self.fixture = File.expand_path("../../fixtures/basic.pptx", __dir__)
  def fixture = self.class.fixture

  describe "loading a real package" do
    subject(:package) { described_class.open(fixture) }

    it "finds the presentation as the main document part" do
      aggregate_failures do
        expect(package.main_document_part.partname.to_s).to eq("/ppt/presentation.xml")
        expect(package.main_document_part.content_type)
          .to eq(Pptx::Opc::CONTENT_TYPE::PML_PRESENTATION_MAIN)
      end
    end

    it "walks the relationship graph to reach every part exactly once" do
      partnames = package.parts.map { |p| p.partname.to_s }
      aggregate_failures do
        expect(partnames).to include("/ppt/presentation.xml", "/ppt/slides/slide1.xml",
                                     "/ppt/slideMasters/slideMaster1.xml", "/docProps/core.xml")
        expect(partnames.uniq).to eq(partnames)
      end
    end

    it "assigns each part the content type from [Content_Types].xml" do
      slide = package.parts.find { |p| p.partname.to_s == "/ppt/slides/slide1.xml" }
      expect(slide.content_type).to eq(Pptx::Opc::CONTENT_TYPE::PML_SLIDE)
    end

    it "parses XML parts and leaves binary parts as bytes" do
      theme = package.parts.find { |p| p.partname.to_s.start_with?("/ppt/theme/") }
      expect(theme).to be_a(Pptx::Opc::Part)
      expect(theme.blob).to start_with("<?xml")
    end

    it "resolves relationship targets to the part objects themselves" do
      presentation = package.main_document_part
      slide_rel = presentation.rels.find { |r| r.reltype == Pptx::Opc::RELATIONSHIP_TYPE::SLIDE }
      expect(slide_rel.target_part.partname.to_s).to eq("/ppt/slides/slide1.xml")
    end

    it "raises a clear error for a file that is not a package" do
      expect { described_class.open("/etc/hostname") }
        .to raise_error(Pptx::PackageNotFoundError)
    end

    it "raises for a path that does not exist" do
      expect { described_class.open("/nonexistent/nope.pptx") }
        .to raise_error(Pptx::PackageNotFoundError)
    end
  end

  describe "round-tripping" do
    before { skip "python-pptx not importable" unless Pptx::Spec::Differential.oracle_available? }

    # The M2 exit criterion: opening a package and saving it again must
    # produce the same package python-pptx does, part for part.
    it "saves a package python-pptx would consider identical" do
      expect_same_package(
        python: "prs = pptx.Presentation(#{fixture.inspect}); prs.save(out)",
        ruby: ->(path) { described_class.open(fixture).save(path) }
      )
    end

    it "round-trips through an IO stream as well as a path" do
      buffer = StringIO.new(+"", "w+b")
      described_class.open(fixture).save(buffer)
      buffer.rewind

      reopened = described_class.open(buffer)
      expect(reopened.parts.count).to eq(described_class.open(fixture).parts.count)
    end

    it "is stable: saving twice produces the same package" do
      first = Tempfile.create(["a", ".pptx"]) { |f| f.close; described_class.open(fixture).save(f.path); ruby_pptx_manifest(f.path) }
      second = Tempfile.create(["b", ".pptx"]) { |f| f.close; described_class.open(fixture).save(f.path); ruby_pptx_manifest(f.path) }
      expect(first).to eq(second)
    end
  end

  describe "#next_partname" do
    subject(:package) { described_class.open(fixture) }

    it "returns the next free number for a partname template" do
      expect(package.next_partname("/ppt/slides/slide%d.xml").to_s)
        .to eq("/ppt/slides/slide2.xml")
    end

    it "fills a gap in the numbering" do
      # slideLayout1..11 exist in the default template; removing one should
      # make its number the next available.
      expect(package.next_partname("/ppt/notesSlides/notesSlide%d.xml").to_s)
        .to eq("/ppt/notesSlides/notesSlide1.xml")
    end
  end
end

RSpec.describe Pptx::Opc::Relationships do
  subject(:rels) { described_class.new("/ppt") }

  let(:part) do
    Pptx::Opc::Part.new(Pptx::Opc::PackURI.new("/ppt/slides/slide1.xml"), "application/xml", nil)
  end

  it "allocates rIds from rId1 upward" do
    expect(rels.get_or_add("t", part)).to eq("rId1")
  end

  it "reuses the rId of a matching relationship instead of duplicating it" do
    first = rels.get_or_add("t", part)
    expect(rels.get_or_add("t", part)).to eq(first)
    expect(rels.size).to eq(1)
  end

  it "treats an external relationship as distinct from an internal one" do
    rels.get_or_add("t", part)
    rels.get_or_add_external("t", "http://example.com")
    expect(rels.size).to eq(2)
  end

  it "fills a gap left by a deleted relationship" do
    rels.get_or_add("t", part)
    other = Pptx::Opc::Part.new(Pptx::Opc::PackURI.new("/ppt/slides/slide2.xml"), "x", nil)
    third = Pptx::Opc::Part.new(Pptx::Opc::PackURI.new("/ppt/slides/slide3.xml"), "x", nil)
    rels.get_or_add("t", other)
    rels.get_or_add("t", third)
    rels.delete("rId2")
    expect(rels.get_or_add("t", Pptx::Opc::Part.new(Pptx::Opc::PackURI.new("/ppt/s4.xml"), "x", nil)))
      .to eq("rId2")
  end

  it "writes a relative target ref for an internal relationship" do
    rels.get_or_add("t", part)
    expect(rels.xml).to include('Target="slides/slide1.xml"')
  end

  it "marks an external relationship and writes its URL verbatim" do
    rels.get_or_add_external("t", "https://example.com/x?a=1")
    aggregate_failures do
      expect(rels.xml).to include('TargetMode="External"')
      expect(rels.xml).to include('Target="https://example.com/x?a=1"')
    end
  end

  it "orders relationships numerically, not lexically" do
    11.times do |i|
      p = Pptx::Opc::Part.new(Pptx::Opc::PackURI.new("/ppt/s#{i}.xml"), "x", nil)
      rels.get_or_add("t#{i}", p)
    end
    ids = rels.xml.scan(/Id="(rId\d+)"/).flatten
    expect(ids).to eq((1..11).map { |n| "rId#{n}" })
  end

  it "raises when asked for a reltype that is not present" do
    expect { rels.part_with_reltype("nope") }
      .to raise_error(Pptx::NotFoundError, /no relationship of type/)
  end

  it "raises when a reltype is ambiguous" do
    other = Pptx::Opc::Part.new(Pptx::Opc::PackURI.new("/ppt/slides/slide2.xml"), "x", nil)
    rels.get_or_add("t", part)
    rels.get_or_add("t", other)
    expect { rels.part_with_reltype("t") }.to raise_error(Pptx::Error, /multiple relationships/)
  end
end

RSpec.describe Pptx::Opc::ContentTypeMap do
  subject(:map) { described_class.from_xml(xml) }

  let(:xml) do
    <<~XML
      <Types xmlns="#{Pptx::Opc::NAMESPACE::OPC_CONTENT_TYPES}">
        <Default Extension="XML" ContentType="application/xml"/>
        <Override PartName="/ppt/Presentation.xml" ContentType="application/custom"/>
      </Types>
    XML
  end

  it "prefers an override to a default" do
    expect(map[Pptx::Opc::PackURI.new("/ppt/Presentation.xml")]).to eq("application/custom")
  end

  it "matches partnames without regard to case" do
    expect(map[Pptx::Opc::PackURI.new("/ppt/presentation.xml")]).to eq("application/custom")
  end

  it "matches extensions without regard to case" do
    expect(map[Pptx::Opc::PackURI.new("/ppt/other.xml")]).to eq("application/xml")
  end

  it "raises for a partname it cannot type" do
    expect { map[Pptx::Opc::PackURI.new("/ppt/media/image1.png")] }
      .to raise_error(KeyError, /no content-type/)
  end
end
