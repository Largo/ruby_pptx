# frozen_string_literal: true

# Sections are a PowerPoint 2010 extension that python-pptx does not model, so
# there is no differential oracle for them. They are checked instead against
# the structure given in the Microsoft [MS-PPTX] specification, and by proving
# python-pptx preserves them across a round-trip.
RSpec.describe Pptx::Sections do
  let(:presentation) { Pptx::Presentation.new_default }
  let!(:slides) { Array.new(3) { presentation.slides.add(presentation.slide_layouts[6]) } }

  subject(:sections) { presentation.sections }

  it "starts empty and writes no extension element" do
    aggregate_failures do
      expect(sections).to be_empty
      expect(presentation.element.section_list).to be_nil
      expect(presentation.element.xml).not_to include("extLst")
    end
  end

  it "adds a named section" do
    section = sections.add("Introduction")
    aggregate_failures do
      expect(section.name).to eq("Introduction")
      expect(sections.size).to eq(1)
      expect(sections.map(&:name)).to eq(["Introduction"])
    end
  end

  it "indexes by position and looks up by name" do
    sections.add("One")
    two = sections.add("Two")
    aggregate_failures do
      expect(sections[1]).to eq(two)
      expect(sections["Two"]).to eq(two)
      expect(sections["Nope"]).to be_nil
    end
  end

  it "puts the given slides in the section" do
    section = sections.add("Intro", slides: [slides[0], slides[1]])
    aggregate_failures do
      expect(section.slide_ids).to eq([slides[0].slide_id, slides[1].slide_id])
      expect(section.slides.map(&:slide_id)).to eq([slides[0].slide_id, slides[1].slide_id])
      expect(section).to include(slides[0])
      expect(section).not_to include(slides[2])
    end
  end

  # PowerPoint gives each slide to exactly one section; a slide appearing twice
  # would make the thumbnail pane ambiguous.
  it "moves a slide out of its previous section rather than duplicating it" do
    first = sections.add("First", slides: [slides[0], slides[1]])
    second = sections.add("Second")
    second << slides[0]
    aggregate_failures do
      expect(first.slide_ids).to eq([slides[1].slide_id])
      expect(second.slide_ids).to eq([slides[0].slide_id])
    end
  end

  it "ignores a slide added to the same section twice" do
    section = sections.add("Intro", slides: [slides[0]])
    section << slides[0]
    expect(section.slide_ids).to eq([slides[0].slide_id])
  end

  it "renames a section" do
    section = sections.add("Old")
    section.name = "New"
    expect(sections.map(&:name)).to eq(["New"])
  end

  it "removes the extension entirely when the last section goes" do
    section = sections.add("Only")
    sections.delete(section)
    aggregate_failures do
      expect(sections).to be_empty
      expect(presentation.element.section_list).to be_nil
      expect(presentation.element.xml).not_to include("sectionLst")
    end
  end

  it "keeps the slides when a section is deleted" do
    section = sections.add("Intro", slides: slides)
    sections.delete(section)
    expect(presentation.slides.size).to eq(3)
  end

  # A section stores slide ids, so a deleted slide leaves an id with nothing
  # behind it rather than a dangling object.
  it "skips a slide id that no longer resolves" do
    section = sections.add("Intro", slides: [slides[0]])
    allow(presentation.slides).to receive(:by_id).and_return(nil)
    expect(section.slides).to be_empty
  end

  it "gives each section a brace-wrapped uppercase GUID" do
    guid = /\A\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}\z/
    expect(sections.add("Intro").id).to match(guid)
  end

  describe "the XML it writes" do
    # Compared against the example in [MS-PPTX] section 2.3.1.25.
    it "matches the structure given in the Microsoft specification" do
      sections.add("Introduction", slides: [slides[0]])
      sections.add("Content", slides: [slides[1], slides[2]])

      xml = presentation.element.xpath("./p:extLst").first.xml.gsub(/\s+/, " ")
      aggregate_failures do
        expect(xml).to include(%(<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">))
        expect(xml).to include(
          %(<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">)
        )
        expect(xml).to match(/<p14:section name="Introduction" id="\{[0-9A-F-]+\}">/)
        expect(xml).to include(%(<p14:sldIdLst> <p14:sldId id="#{slides[1].slide_id}"/>))
      end
    end

    it "places the extension list last in the presentation element" do
      sections.add("Intro")
      expect(presentation.element.element_children.last.nsptag).to eq("p:extLst")
    end
  end

  describe "round-tripping through python-pptx" do
    before { require_oracle! }

    # python-pptx does not know what sections are, but it must not destroy
    # them: it carries the presentation part through untouched.
    it "survives being opened and saved by python-pptx" do
      sections.add("Introduction", slides: [slides[0]])
      sections.add("Content", slides: [slides[1], slides[2]])

      Tempfile.create(["sections", ".pptx"]) do |file|
        file.close
        presentation.save(file.path)

        script = <<~PY
          import sys, zipfile, re
          import pptx
          prs = pptx.Presentation(#{file.path.inspect})
          prs.save(#{file.path.inspect})
          xml = zipfile.ZipFile(#{file.path.inspect}).read("ppt/presentation.xml").decode()
          m = re.search(r"(?s)<p14:sectionLst.*</p14:sectionLst>", xml)
          sys.stdout.write(m.group(0) if m else "")
        PY
        out, err, status = Open3.capture3("python3", "-c", script)
        raise "round-trip failed: #{err}" unless status.success?

        aggregate_failures do
          expect(out).to include('name="Introduction"', 'name="Content"')
          expect(out.scan("<p14:sldId ").size).to eq(3)
        end

        reopened = Pptx::Presentation.open(file.path)
        expect(reopened.sections.map(&:name)).to eq(%w[Introduction Content])
      end
    end
  end
end
