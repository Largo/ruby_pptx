# frozen_string_literal: true

RSpec.describe "speaker notes" do
  let(:deck) { Pptx::Presentation.new_default }
  let(:slide) { deck.slides.add(deck.slide_layouts[1]) }

  describe "reading and writing" do
    it "reports no notes, and creates nothing, until some are set" do
      aggregate_failures do
        expect(slide.notes).to be_nil
        expect(slide).not_to be_notes_slide
        partnames = deck.part.package.parts.map { |p| p.partname.to_s }
        expect(partnames.grep(/notes/)).to eq([])
      end
    end

    it "round-trips the notes text" do
      slide.notes = "Mention the Q3 numbers"
      aggregate_failures do
        expect(slide).to be_notes_slide
        expect(slide.notes).to eq("Mention the Q3 numbers")
      end
    end

    it "keeps paragraphs as the text frame does" do
      slide.notes = "First point\nSecond point"
      expect(slide.notes_slide.notes_text_frame.paragraphs.size).to eq(2)
    end

    it "returns the same notes slide every time" do
      expect(slide.notes_slide.element).to eq(slide.notes_slide.element)
    end

    it "creates one notes master however many slides get notes" do
      3.times { deck.slides.add(deck.slide_layouts[6]).notes = "x" }
      partnames = deck.part.package.parts.map { |p| p.partname.to_s }
      aggregate_failures do
        expect(partnames.grep(%r{notesMasters/})).to eq(["/ppt/notesMasters/notesMaster1.xml"])
        expect(partnames.grep(%r{notesSlides/}).size).to eq(3)
      end
    end

    it "gives the notes master a theme of its own" do
      deck.notes_master
      partnames = deck.part.package.parts.map { |p| p.partname.to_s }
      expect(partnames.grep(%r{theme/}).sort).to eq(%w[/ppt/theme/theme1.xml /ppt/theme/theme2.xml])
    end

    it "survives a save and reload" do
      slide.notes = "Keep this"
      reloaded = Tempfile.create(["notes", ".pptx"]) do |file|
        file.close
        deck.save(file.path)
        Pptx::Presentation.open(file.path)
      end
      aggregate_failures do
        expect(reloaded.slides[0].notes).to eq("Keep this")
        expect(reloaded.slides[0].notes_slide).to be_a(Pptx::NotesSlide)
      end
    end
  end

  describe Pptx::NotesSlide do
    subject(:notes_slide) { slide.notes_slide }

    # Header, date and footer stay on the master, as PowerPoint does it.
    it "clones only the slide image, body and slide number from the master" do
      types = notes_slide.placeholders.map { |ph| ph.placeholder_format.type.name }
      expect(types).to eq(%i[SLIDE_IMAGE BODY SLIDE_NUMBER])
    end

    it "names the placeholders as PowerPoint names notes placeholders" do
      expect(notes_slide.placeholders.map(&:name))
        .to eq(["Slide Image Placeholder 1", "Notes Placeholder 2", "Slide Number Placeholder 3"])
    end

    it "finds the body placeholder as the notes placeholder" do
      expect(notes_slide.notes_placeholder.placeholder_format.type.name).to eq(:BODY)
    end

    it "reports nil rather than raising when the body placeholder is gone" do
      body = notes_slide.notes_placeholder.element
      body.parent.remove(body)
      aggregate_failures do
        expect(notes_slide.notes_placeholder).to be_nil
        expect(notes_slide.notes_text_frame).to be_nil
      end
    end

    it "refuses to set notes text with nowhere to put it" do
      body = notes_slide.notes_placeholder.element
      body.parent.remove(body)
      expect { slide.notes = "lost" }.to raise_error(Pptx::Error, /no notes placeholder/)
    end

    # The placeholders carry no geometry; it comes from the notes master.
    it "inherits placeholder geometry from the notes master" do
      body = notes_slide.notes_placeholder
      master_body = deck.notes_master.placeholders.by_type(body.element.ph_type)
      aggregate_failures do
        expect(body.element.xpath("./p:spPr/a:xfrm")).to be_empty
        expect(body.width).to eq(master_body.width)
        expect(body.top).to eq(master_body.top)
      end
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "adds notes to a slide identically" do
      expect_same_package(
        python: <<~PY,
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[1])
          slide.notes_slide.notes_text_frame.text = "Mention the Q3 numbers"
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[1])
          slide.notes_slide.notes_text_frame.text = "Mention the Q3 numbers"
          prs.save(path)
        }
      )
    end

    # Several slides exercise the part numbering and that the master and its
    # theme are created exactly once; skipping one slide checks that notes
    # are only added where asked for.
    it "adds notes to several slides identically" do
      expect_same_package(
        python: <<~PY,
          prs = pptx.Presentation()
          for i in range(4):
              slide = prs.slides.add_slide(prs.slide_layouts[i])
              if i != 2:
                  slide.notes_slide.notes_text_frame.text = "Note %d\\nsecond line" % i
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          4.times do |i|
            slide = prs.slides.add(prs.slide_layouts[i])
            slide.notes = "Note #{i}\nsecond line" unless i == 2
          end
          prs.save(path)
        }
      )
    end

    it "creates the notes master on its own identically" do
      expect_same_package(
        python: <<~PY,
          prs = pptx.Presentation()
          prs.notes_master
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          prs.notes_master
          prs.save(path)
        }
      )
    end

    it "reads notes written by python-pptx" do
      text = Tempfile.create(["py-notes", ".pptx"]) do |file|
        file.close
        script = <<~PY
          import sys, pptx
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          slide.notes_slide.notes_text_frame.text = "Written in Python"
          prs.save(sys.argv[1])
        PY
        _, err, status = Open3.capture3("python3", "-c", script, file.path)
        raise "python-pptx failed: #{err}" unless status.success?

        Pptx::Presentation.open(file.path).slides[0].notes
      end
      expect(text).to eq("Written in Python")
    end
  end
end
