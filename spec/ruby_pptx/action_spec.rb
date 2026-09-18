# frozen_string_literal: true

RSpec.describe Pptx::ActionSetting do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }
  let(:other_slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }

  let(:shape) do
    slide.shapes.add_textbox(at: [Pptx.inches(1), Pptx.inches(1)],
                             size: [Pptx.inches(4), Pptx.inches(1)])
  end

  describe "a shape's hyperlink" do
    it "is nil until one is set" do
      aggregate_failures do
        expect(shape.hyperlink).to be_nil
        expect(shape.click_action.action).to eq(Pptx::Enum::PP_ACTION::NONE)
      end
    end

    it "round-trips a URL" do
      shape.hyperlink = "https://example.com/docs?a=1"
      aggregate_failures do
        expect(shape.hyperlink).to eq("https://example.com/docs?a=1")
        expect(shape.click_action.action).to eq(Pptx::Enum::PP_ACTION::HYPERLINK)
      end
    end

    it "stores the target as an external relationship, not in the slide XML" do
      shape.hyperlink = "https://example.com"
      aggregate_failures do
        expect(shape.element.xml).not_to include("example.com")
        expect(shape.element.xml).to match(/<a:hlinkClick[^>]*r:id="rId\d+"/)
        expect(slide.part.rels.any?(&:external?)).to be(true)
      end
    end

    it "replaces rather than accumulating when set twice" do
      shape.hyperlink = "https://first.example"
      shape.hyperlink = "https://second.example"
      aggregate_failures do
        expect(shape.hyperlink).to eq("https://second.example")
        expect(shape.element.xml.scan("<a:hlinkClick").size).to eq(1)
        expect(slide.part.rels.count(&:external?)).to eq(1)
      end
    end

    # A dangling relationship is the kind of thing PowerPoint complains about
    # on open, so removing the link has to remove it too.
    it "drops the relationship when the link is removed" do
      shape.hyperlink = "https://example.com"
      expect { shape.hyperlink = nil }
        .to change { slide.part.rels.count(&:external?) }.from(1).to(0)
      aggregate_failures do
        expect(shape.hyperlink).to be_nil
        expect(shape.element.xml).not_to include("hlinkClick")
      end
    end

    it "treats an empty string as removal" do
      shape.hyperlink = "https://example.com"
      shape.hyperlink = ""
      expect(shape.hyperlink).to be_nil
    end
  end

  describe "a run's hyperlink" do
    subject(:run) do
      shape.text_frame.text = "Visit the site"
      shape.text_frame.paragraphs.first.runs.first
    end

    it "links just that run, leaving the shape alone" do
      run.hyperlink = "https://example.com/run"
      aggregate_failures do
        expect(run.hyperlink).to eq("https://example.com/run")
        expect(shape.hyperlink).to be_nil
        expect(run.element.xml).to include("hlinkClick")
      end
    end

    it "removes cleanly" do
      run.hyperlink = "https://example.com"
      run.hyperlink = nil
      expect(run.hyperlink).to be_nil
    end
  end

  describe "jumping to a slide" do
    it "records the jump as an internal relationship" do
      shape.click_action.target_slide = other_slide
      aggregate_failures do
        expect(shape.click_action.action).to eq(Pptx::Enum::PP_ACTION::NAMED_SLIDE)
        expect(shape.click_action.target_slide.slide_id).to eq(other_slide.slide_id)
        expect(shape.element.xml).to include('action="ppaction://hlinksldjump"')
      end
    end

    it "is nil when the click is an ordinary hyperlink" do
      shape.hyperlink = "https://example.com"
      expect(shape.click_action.target_slide).to be_nil
    end

    it "clears when assigned nil" do
      shape.click_action.target_slide = other_slide
      shape.click_action.target_slide = nil
      aggregate_failures do
        expect(shape.click_action.action).to eq(Pptx::Enum::PP_ACTION::NONE)
        expect(shape.click_action.target_slide).to be_nil
      end
    end
  end

  describe "reading actions PowerPoint writes" do
    def action_for(attributes)
      shape.element.nvXxPr.cNvPr.get_or_add_hlinkClick.tap do |link|
        attributes.each { |name, value| link.set(name, value) }
      end
      shape.click_action.action
    end

    {
      "ppaction://hlinkshowjump?jump=nextslide" => :NEXT_SLIDE,
      "ppaction://hlinkshowjump?jump=previousslide" => :PREVIOUS_SLIDE,
      "ppaction://hlinkshowjump?jump=firstslide" => :FIRST_SLIDE,
      "ppaction://hlinkshowjump?jump=lastslide" => :LAST_SLIDE,
      "ppaction://hlinkshowjump?jump=endshow" => :END_SHOW,
      "ppaction://hlinkfile" => :OPEN_FILE,
      "ppaction://macro" => :RUN_MACRO,
      "ppaction://program" => :RUN_PROGRAM,
      "ppaction://customshow?id=0&return=true" => :NAMED_SLIDE_SHOW,
      "ppaction://ole?verb=0" => :OLE_VERB
    }.each do |url, expected|
      it "reads #{url} as #{expected}" do
        expect(action_for("action" => url)).to eq(Pptx::Enum::PP_ACTION.fetch(expected))
      end
    end

    # An action this library does not know about should read as NONE rather
    # than raising: the file is still valid, we just cannot name what it does.
    it "reads an unrecognized action as NONE" do
      expect(action_for("action" => "ppaction://somethingnew"))
        .to eq(Pptx::Enum::PP_ACTION::NONE)
    end

    # python-pptx's `address` returns the relationship target for any action
    # carrying one, so a slide jump reports the target slide's partname. That
    # is matched here; `url` is the accessor that means "a hyperlink", and it
    # is what `shape.hyperlink` uses.
    it "reports the relationship target for a slide jump, as python-pptx does" do
      shape.click_action.target_slide = other_slide
      aggregate_failures do
        expect(shape.click_action.address).to eq("slide2.xml")
        expect(shape.click_action.url).to be_nil
        expect(shape.hyperlink).to be_nil
      end
    end

    it "reports no target at all for an action with no relationship" do
      shape.element.nvXxPr.cNvPr.get_or_add_hlinkClick.set("action", "ppaction://macro")
      aggregate_failures do
        expect(shape.click_action.action).to eq(Pptx::Enum::PP_ACTION::RUN_MACRO)
        expect(shape.click_action.address).to be_nil
        expect(shape.hyperlink).to be_nil
      end
    end
  end

  describe Pptx::Oxml::CT_Hyperlink do
    subject(:link) { shape.element.nvXxPr.cNvPr.get_or_add_hlinkClick }

    it "splits the action URL into a verb and its fields" do
      link.set("action", "ppaction://customshow?id=0&return=true")
      aggregate_failures do
        expect(link.action_verb).to eq("customshow")
        expect(link.action_fields).to eq("id" => "0", "return" => "true")
      end
    end

    it "reports no verb or fields for a plain hyperlink" do
      aggregate_failures do
        expect(link.action_verb).to be_nil
        expect(link.action_fields).to eq({})
      end
    end

    it "reports no fields when the action has no query" do
      link.set("action", "ppaction://macro")
      expect(link.action_fields).to eq({})
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "writes hyperlinks and a slide jump identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          prs = pptx.Presentation()
          s1 = prs.slides.add_slide(prs.slide_layouts[6])
          s2 = prs.slides.add_slide(prs.slide_layouts[6])

          box = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
          box.text_frame.text = "Visit the site"
          box.click_action.hyperlink.address = "https://example.com/docs?a=1"
          box.text_frame.paragraphs[0].runs[0].hyperlink.address = "https://example.com/run"

          from pptx.enum.shapes import MSO_SHAPE
          jump = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(1), Inches(3), Inches(2), Inches(1))
          jump.click_action.target_slide = s2
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          s1 = prs.slides.add(prs.slide_layouts[6])
          s2 = prs.slides.add(prs.slide_layouts[6])

          box = s1.shapes.add_textbox(at: [Pptx.inches(1), Pptx.inches(1)],
                                      size: [Pptx.inches(4), Pptx.inches(1)])
          box.text_frame.text = "Visit the site"
          box.hyperlink = "https://example.com/docs?a=1"
          box.text_frame.paragraphs.first.runs.first.hyperlink = "https://example.com/run"

          jump = s1.shapes.add_shape(:rounded_rectangle,
                                     at: [Pptx.inches(1), Pptx.inches(3)],
                                     size: [Pptx.inches(2), Pptx.inches(1)])
          jump.click_action.target_slide = s2
          prs.save(path)
        }
      )
    end
  end
end
