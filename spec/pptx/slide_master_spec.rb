# frozen_string_literal: true

require "json"
require "open3"

RSpec.describe "defining a slide master in code" do
  let(:deck) { Pptx::Presentation.new_default }

  # The master the default template already has, so "the new one" is always
  # unambiguous.
  def existing_master_count = 1

  describe Pptx::SlideMasters do
    subject(:master) { deck.slide_masters.add(name: "Corporate") }

    it "appends a master to the presentation" do
      master
      expect(deck.slide_masters.size).to eq(existing_master_count + 1)
    end

    it "creates a slide master part and a theme part of its own" do
      master
      partnames = deck.part.package.parts.map { |p| p.partname.to_s }
      aggregate_failures do
        expect(partnames).to include("/ppt/slideMasters/slideMaster2.xml")
        expect(partnames).to include("/ppt/theme/theme2.xml")
      end
    end

    it "relates the master to its own theme, not the existing one" do
      theme_partname = master.part.theme_part.partname.to_s
      original = deck.slide_masters[0].part.theme_part.partname.to_s
      aggregate_failures do
        expect(theme_partname).to eq("/ppt/theme/theme2.xml")
        expect(theme_partname).not_to eq(original)
      end
    end

    # PowerPoint numbers masters from 2147483648 upwards, above the slide-id
    # range, so the two can never be confused.
    it "numbers the master id above the slide id range" do
      master
      ids = deck.part.element.sldMasterIdLst.sldMasterId_list.map(&:id)
      aggregate_failures do
        expect(ids.size).to eq(2)
        expect(ids).to all(be >= 2_147_483_648)
        expect(ids.uniq.size).to eq(2)
        expect(ids.last).to be > ids.first
      end
    end

    it "yields the master when given a block" do
      yielded = nil
      returned = deck.slide_masters.add(name: "X") { |m| yielded = m }
      expect(yielded).to eq(returned)
    end

    it "rejects an unknown placeholders option" do
      expect { deck.slide_masters.add(placeholders: :some) }
        .to raise_error(ArgumentError, /:standard or :none/)
    end

    it "can create a master with no placeholders at all" do
      bare = deck.slide_masters.add(name: "Bare", placeholders: :none)
      expect(bare.placeholders.size).to eq(0)
    end

    describe "the standard placeholder set" do
      it "has the five placeholders PowerPoint expects" do
        types = master.placeholders.map { |ph| ph.element.ph_type.name }
        expect(types).to eq(%i[TITLE BODY DATE FOOTER SLIDE_NUMBER])
      end

      it "gives each non-title placeholder a distinct idx" do
        indexes = master.placeholders.map { |ph| ph.element.ph_idx }
        expect(indexes).to eq([0, 1, 2, 3, 4])
      end

      # The fractions come from the Office default master, whose slide is
      # exactly this size, so a default-sized deck must reproduce its geometry
      # to the EMU rather than merely landing close.
      it "reproduces the Office geometry exactly at the default slide size" do
        geometry = master.placeholders.map { |ph| [ph.left.emu, ph.top.emu, ph.width.emu, ph.height.emu] }
        expect(geometry).to eq(
          [[457_200, 274_638, 8_229_600, 1_143_000],
           [457_200, 1_600_200, 8_229_600, 4_525_963],
           [457_200, 6_356_350, 2_133_600, 365_125],
           [3_124_200, 6_356_350, 2_895_600, 365_125],
           [6_553_200, 6_356_350, 2_133_600, 365_125]]
        )
      end

      it "scales the geometry to a widescreen slide" do
        deck.slide_width = Pptx.inches(13.333)
        wide = deck.slide_masters.add(name: "Wide")
        title = wide.placeholders.first
        narrow_title = Pptx::Presentation.new_default.slide_masters.add.placeholders.first
        aggregate_failures do
          expect(title.width).to be > narrow_title.width
          # Same fraction of the slide, so the same proportion either way.
          expect(title.width.emu.to_f / deck.slide_width.emu)
            .to be_within(0.0001).of(narrow_title.width.emu.to_f / Pptx.inches(10).emu)
        end
      end
    end
  end

  describe Pptx::Theme do
    subject(:theme) { deck.slide_masters.add(name: "Corporate").theme }

    it "names the theme and both schemes together" do
      aggregate_failures do
        expect(theme.name).to eq("Corporate")
        expect(theme.element.clrScheme.name).to eq("Corporate")
        expect(theme.element.fontScheme.name).to eq("Corporate")
      end
    end

    it "reads and writes each of the twelve theme colours" do
      theme.colors[:accent1].rgb = Pptx::RGBColor.from_string("1F497D")
      aggregate_failures do
        expect(theme.colors.names.size).to eq(12)
        expect(theme.colors[:accent1].rgb.to_s).to eq("1F497D")
      end
    end

    it "sets several colours at once from hex strings" do
      theme.colors.update(accent1: "1F497D", accent2: "C0504D")
      aggregate_failures do
        expect(theme.colors[:accent1].rgb.to_s).to eq("1F497D")
        expect(theme.colors[:accent2].rgb.to_s).to eq("C0504D")
      end
    end

    it "replaces a system colour with an explicit one" do
      # dk1 arrives as a:sysClr in the base theme.
      theme.colors[:dk1].rgb = Pptx::RGBColor.from_string("112233")
      expect(theme.colors[:dk1].rgb.to_s).to eq("112233")
    end

    it "refuses a colour name that is not in the scheme" do
      expect { theme.colors[:accent7] }.to raise_error(Pptx::NotFoundError, /accent7/)
    end

    it "enumerates the colours in PowerPoint's order" do
      expect(theme.colors.map { |name, _| name }.first(5))
        .to eq(%i[dk1 lt1 dk2 lt2 accent1])
    end

    it "reads and writes the heading and body typefaces" do
      theme.fonts.major = "Georgia"
      theme.fonts.minor = "Verdana"
      aggregate_failures do
        expect(theme.fonts.major).to eq("Georgia")
        expect(theme.fonts.minor).to eq("Verdana")
      end
    end

    # Two masters must not end up editing one another's colours.
    it "gives each master an independent theme" do
      other = deck.slide_masters.add(name: "Other")
      theme.colors[:accent1].rgb = Pptx::RGBColor.from_string("1F497D")
      other.theme.colors[:accent1].rgb = Pptx::RGBColor.from_string("C0504D")
      aggregate_failures do
        expect(theme.colors[:accent1].rgb.to_s).to eq("1F497D")
        expect(other.theme.colors[:accent1].rgb.to_s).to eq("C0504D")
      end
    end
  end

  describe Pptx::SlideLayouts do
    let(:master) { deck.slide_masters.add(name: "Corporate") }

    it "adds a layout to the master" do
      layout = master.slide_layouts.add("Title and Content", type: "obj")
      aggregate_failures do
        expect(master.slide_layouts.size).to eq(1)
        expect(layout.name).to eq("Title and Content")
        expect(layout.type).to eq("obj")
      end
    end

    it "creates a layout part related back to the master" do
      layout = master.slide_layouts.add("Blank", type: "blank")
      aggregate_failures do
        expect(layout.part.partname.to_s).to match(%r{/ppt/slideLayouts/slideLayout\d+\.xml})
        expect(layout.slide_master).to eq(master)
      end
    end

    # Without preserve="1" PowerPoint discards a layout that no slide uses.
    it "marks the layout preserved" do
      layout = master.slide_layouts.add("Blank")
      expect(layout.element.preserve).to be(true)
    end

    it "numbers layout ids above the slide id range and keeps them distinct" do
      3.times { |i| master.slide_layouts.add("L#{i}") }
      ids = master.element.sldLayoutIdLst.sldLayoutId_list.map(&:id)
      aggregate_failures do
        expect(ids).to all(be >= 2_147_483_649)
        expect(ids.uniq.size).to eq(3)
      end
    end

    it "defaults to the obj layout type" do
      expect(master.slide_layouts.add("Content").type).to eq("obj")
    end

    it "accepts a symbol layout type" do
      expect(master.slide_layouts.add("Section", type: :secHead).type).to eq("secHead")
    end

    it "refuses a layout type the schema does not define" do
      expect { master.slide_layouts.add("Nope", type: "wibble") }
        .to raise_error(ArgumentError, /not a slide layout type/)
    end

    it "can leave the layout type unset" do
      expect(master.slide_layouts.add("Untyped", type: nil).type).to be_nil
    end

    it "yields the layout when given a block" do
      yielded = nil
      returned = master.slide_layouts.add("L") { |l| yielded = l }
      expect(yielded).to eq(returned)
    end

    it "finds the new layout by name on the master" do
      master.slide_layouts.add("Section Header", type: "secHead")
      expect(master.slide_layouts.by_name("Section Header")).not_to be_nil
    end
  end

  describe Pptx::PlaceholderAuthoring do
    let(:master) { deck.slide_masters.add(name: "Corporate", placeholders: :none) }
    let(:layout) { master.slide_layouts.add("Custom") }

    it "adds a placeholder with explicit geometry" do
      ph = layout.placeholders.add(:title, at: [Pptx.inches(1), Pptx.inches(2)],
                                           size: [Pptx.inches(8), Pptx.inches(1)])
      aggregate_failures do
        expect(ph.left).to eq(Pptx.inches(1))
        expect(ph.top).to eq(Pptx.inches(2))
        expect(ph.width).to eq(Pptx.inches(8))
        expect(ph.height).to eq(Pptx.inches(1))
      end
    end

    # A placeholder with no xfrm inherits its position, which is the whole
    # point of one -- so omitting the geometry must not invent any.
    it "leaves geometry inherited when it is not given" do
      ph = layout.placeholders.add(:body, idx: 1)
      expect(ph.element.xpath("./p:spPr/a:xfrm")).to be_empty
    end

    # PowerPoint numbers a placeholder by its position in the shape tree, so
    # the body added after a title is "2".
    it "names the placeholder the way PowerPoint does" do
      layout.placeholders.add(:title)
      aggregate_failures do
        expect(layout.placeholders.first.name).to eq("Title 1")
        expect(layout.placeholders.add(:body, idx: 1).name).to eq("Text Placeholder 2")
      end
    end

    # There is only ever one title and PowerPoint finds it by type.
    it "gives a title no idx" do
      layout.placeholders.add(:title)
      expect(layout.placeholders.first.element.ph_idx).to eq(0)
    end

    it "allocates the next free idx when none is given" do
      layout.placeholders.add(:title)
      layout.placeholders.add(:body)
      layout.placeholders.add(:body)
      indexes = layout.placeholders.map { |ph| ph.element.ph_idx }
      expect(indexes).to eq([0, 1, 2])
    end

    it "works around an idx already taken" do
      layout.placeholders.add(:body, idx: 1)
      layout.placeholders.add(:body, idx: 3)
      expect(layout.placeholders.add(:body).element.ph_idx).to eq(2)
    end

    it "gives each placeholder a distinct shape id" do
      3.times { layout.placeholders.add(:body) }
      ids = layout.placeholders.map(&:shape_id)
      expect(ids.uniq.size).to eq(3)
    end

    it "records the size hint and the orientation" do
      ph = layout.placeholders.add(:body, idx: 1, sz: :quarter, orient: :vertical)
      aggregate_failures do
        expect(ph.element.ph_sz).to eq("quarter")
        expect(ph.element.ph_orient).to eq("vert")
      end
    end

    it "refuses a size hint that is not one of the three" do
      expect { layout.placeholders.add(:body, sz: :massive) }
        .to raise_error(ArgumentError, /:full, :half or :quarter/)
    end

    it "refuses an orientation it does not understand" do
      expect { layout.placeholders.add(:body, orient: :sideways) }
        .to raise_error(ArgumentError, /:vertical or nil/)
    end

    it "refuses a placeholder type that does not exist" do
      expect { layout.placeholders.add(:wibble) }
        .to raise_error(ArgumentError, /not a member of PP_PLACEHOLDER_TYPE/)
    end

    it "adds placeholders to a master as well as a layout" do
      master.placeholders.add(:title, at: [0, 0], size: [Pptx.inches(9), Pptx.inches(1)])
      expect(master.placeholders.size).to eq(1)
    end
  end

  describe "a slide built on a generated layout" do
    let(:master) { deck.slide_masters.add(name: "Corporate") }
    let(:layout) do
      master.slide_layouts.add("Title and Content", type: "obj") do |l|
        l.placeholders.add(:title, at: [Pptx.inches(0.5), Pptx.inches(0.3)],
                                   size: [Pptx.inches(9), Pptx.inches(1.25)])
        l.placeholders.add(:body, idx: 1, at: [Pptx.inches(0.5), Pptx.inches(1.75)],
                                  size: [Pptx.inches(9), Pptx.inches(4.5)])
      end
    end

    it "clones the layout's placeholders onto the slide" do
      slide = deck.slides.add(layout)
      expect(slide.placeholders.map { |ph| ph.element.ph_idx }).to eq([0, 1])
    end

    it "inherits the layout's geometry" do
      slide = deck.slides.add(layout)
      expect(slide.shapes.title.width).to eq(Pptx.inches(9))
    end

    it "takes text like any other slide" do
      slide = deck.slides.add(layout)
      slide.shapes.title.text = "Hello"
      slide.placeholders[1].text_frame.text = "Alpha\nBeta"
      aggregate_failures do
        expect(slide.shapes.title.text).to eq("Hello")
        expect(slide.placeholders[1].text_frame.text).to eq("Alpha\nBeta")
      end
    end

    it "survives a save and reload" do
      slide = deck.slides.add(layout)
      slide.shapes.title.text = "Hello"
      reloaded = Tempfile.create(["master", ".pptx"]) do |file|
        file.close
        deck.save(file.path)
        Pptx::Presentation.open(file.path)
      end
      aggregate_failures do
        expect(reloaded.slide_masters.size).to eq(2)
        expect(reloaded.slides[0].shapes.title.text).to eq("Hello")
        expect(reloaded.slide_masters[1].theme.name).to eq("Corporate")
      end
    end
  end

  # A placeholder normally carries no geometry of its own, so reading it back
  # has to walk slide -> layout -> master. Without this a freshly built slide
  # looks like it has no positions at all.
  describe Pptx::InheritsDimensions do
    it "reads a slide placeholder's geometry from its layout" do
      slide = deck.slides.add(deck.slide_layouts[0])
      layout_title = deck.slide_layouts[0].placeholders.by_idx(0)
      aggregate_failures do
        expect(slide.shapes.title.element.xpath("./p:spPr/a:xfrm")).to be_empty
        expect(slide.shapes.title.width).to eq(layout_title.width)
        expect(slide.shapes.title.left).to eq(layout_title.left)
      end
    end

    it "stops inheriting once a value is set directly" do
      slide = deck.slides.add(deck.slide_layouts[0])
      title = slide.shapes.title
      title.width = Pptx.inches(3)
      aggregate_failures do
        expect(title.width).to eq(Pptx.inches(3))
        # The others still come from the layout.
        expect(title.left).to eq(deck.slide_layouts[0].placeholders.by_idx(0).left)
      end
    end

    # A layout placeholder inherits from the master placeholder of the
    # corresponding type, not the same one -- a body, chart and table
    # placeholder all take their geometry from the master's body.
    it "reads a layout placeholder's geometry from the master" do
      master = deck.slide_masters.add(name: "Corporate")
      layout = master.slide_layouts.add("Custom")
      body = layout.placeholders.add(:table, idx: 1)
      aggregate_failures do
        expect(body.element.xpath("./p:spPr/a:xfrm")).to be_empty
        expect(body.width).to eq(master.placeholders.by_type(Pptx::Enum::PP_PLACEHOLDER_TYPE::BODY).width)
      end
    end

    it "reports nil when there is nothing to inherit from" do
      master = deck.slide_masters.add(name: "Bare", placeholders: :none)
      layout = master.slide_layouts.add("Custom")
      expect(layout.placeholders.add(:body, idx: 1).width).to be_nil
    end

    it "leaves a master placeholder's own geometry alone" do
      master = deck.slide_masters.add(name: "Corporate")
      expect(master.placeholders.first.width).to eq(Pptx.inches(9))
    end

    # The values must be the ones python-pptx resolves, not merely non-nil.
    it "resolves the same geometry python-pptx does" do
      require_oracle!
      slide = deck.slides.add(deck.slide_layouts[0])
      ours = slide.placeholders.map { |ph| [ph.left, ph.top, ph.width, ph.height].map { |v| v&.emu } }
      script = <<~PY
        import json, sys, pptx
        p = pptx.Presentation()
        s = p.slides.add_slide(p.slide_layouts[0])
        json.dump([[ph.left, ph.top, ph.width, ph.height] for ph in s.placeholders], sys.stdout)
      PY
      out, err, status = Open3.capture3("python3", "-c", script)
      raise "placeholder oracle failed: #{err}" unless status.success?

      expect(ours).to eq(JSON.parse(out))
    end
  end

  # python-pptx cannot create a master, so this is not a differential: it is a
  # round-trip. Our package is opened by the reference implementation and every
  # structural claim is read back from it.
  describe "what python-pptx sees in a generated master" do
    before { require_oracle! }

    def oracle(path)
      script = File.expand_path("../../tools/master_oracle.py", __dir__)
      out, err, status = Open3.capture3("python3", script, path)
      raise "master oracle failed: #{err}" unless status.success?

      JSON.parse(out)
    end

    def build_and_read
      master = deck.slide_masters.add(name: "Corporate")
      master.theme.colors.update(accent1: "1F497D", accent2: "C0504D")
      master.theme.fonts.major = "Georgia"
      master.theme.fonts.minor = "Verdana"
      master.slide_layouts.add("Title and Content", type: "obj") do |layout|
        layout.placeholders.add(:title, at: [Pptx.inches(0.5), Pptx.inches(0.3)],
                                        size: [Pptx.inches(9), Pptx.inches(1.25)])
        layout.placeholders.add(:body, idx: 1, at: [Pptx.inches(0.5), Pptx.inches(1.75)],
                                       size: [Pptx.inches(9), Pptx.inches(4.5)])
        slide = deck.slides.add(layout)
        slide.shapes.title.text = "Built from a hand-made master"
        slide.placeholders[1].text_frame.text = "Alpha\nBeta"
      end
      Tempfile.create(["master", ".pptx"]) do |file|
        file.close
        deck.save(file.path)
        oracle(file.path)
      end
    end

    let(:result) { build_and_read }
    let(:generated) { result["masters"].last }

    it "sees the generated master and its layout" do
      aggregate_failures do
        expect(result["masters"].size).to eq(2)
        expect(generated["partname"]).to eq("/ppt/slideMasters/slideMaster2.xml")
        expect(generated["layouts"].map { |l| l["name"] }).to eq(["Title and Content"])
        expect(generated["layouts"].first["type"]).to eq("obj")
        expect(generated["layouts"].first["preserve"]).to eq("1")
      end
    end

    it "resolves the master's own theme, colours and fonts" do
      theme = generated["theme"]
      aggregate_failures do
        expect(theme["theme_name"]).to eq("Corporate")
        expect(theme["scheme_name"]).to eq("Corporate")
        expect(theme["colors"]["accent1"]).to eq("1F497D")
        expect(theme["colors"]["accent2"]).to eq("C0504D")
        expect(theme["major"]).to eq("Georgia")
        expect(theme["minor"]).to eq("Verdana")
      end
    end

    it "leaves the original master's theme untouched" do
      original = result["masters"].first["theme"]
      aggregate_failures do
        expect(original["theme_name"]).to eq("Office Theme")
        expect(original["colors"]["accent1"]).not_to eq("1F497D")
      end
    end

    it "sees the master's five placeholders with their geometry" do
      types = generated["placeholders"].map { |ph| ph["type"] }
      aggregate_failures do
        expect(types).to eq(%w[TITLE BODY DATE FOOTER SLIDE_NUMBER])
        expect(generated["placeholders"].first.values_at("left", "top", "width", "height"))
          .to eq([457_200, 274_638, 8_229_600, 1_143_000])
      end
    end

    it "sees the layout's placeholders at the positions we set" do
      placeholders = generated["layouts"].first["placeholders"]
      aggregate_failures do
        expect(placeholders.map { |ph| ph["idx"] }).to eq([0, 1])
        expect(placeholders.first["width"]).to eq(Pptx.inches(9).emu)
        expect(placeholders.last["top"]).to eq(Pptx.inches(1.75).emu)
      end
    end

    # The point of the whole exercise: a slide really does inherit from the
    # master we built, not from the one the template came with.
    it "sees the slide inheriting from the generated master" do
      slide = result["slides"].first
      aggregate_failures do
        expect(slide["layout"]).to eq("Title and Content")
        expect(slide["master"]).to eq("/ppt/slideMasters/slideMaster2.xml")
        expect(slide["texts"]).to include("Built from a hand-made master")
      end
    end

    it "sees the slide's placeholders inheriting the layout's width" do
      slide = result["slides"].first
      expect(slide["placeholders"].first["width"]).to eq(Pptx.inches(9).emu)
    end
  end

  # The standard itself, rather than another implementation of it. There is no
  # oracle that can create a master to diff against, so this is what pins the
  # generated XML to something outside our own opinion of it.
  describe "conformance with ISO/IEC 29500-4" do
    before { require_schemas! }

    def save_generated
      master = deck.slide_masters.add(name: "Corporate")
      master.theme.colors.update(accent1: "1F497D")
      master.theme.fonts.major = "Georgia"
      layout = master.slide_layouts.add("Title and Content", type: "obj")
      layout.placeholders.add(:title, at: [Pptx.inches(0.5), Pptx.inches(0.3)],
                                      size: [Pptx.inches(9), Pptx.inches(1.25)])
      layout.placeholders.add(:body, idx: 1, sz: :half,
                                     at: [Pptx.inches(0.5), Pptx.inches(1.75)],
                                     size: [Pptx.inches(9), Pptx.inches(4.5)])
      slide = deck.slides.add(layout)
      slide.shapes.title.text = "Hello"
      slide
    end

    it "produces a package whose every part validates" do
      save_generated
      violations = Tempfile.create(["master", ".pptx"]) do |file|
        file.close
        deck.save(file.path)
        schema_violations(file.path)
      end
      expect(violations).to eq({})
    end

    it "validates a bare master with no placeholders and no layouts" do
      deck.slide_masters.add(name: "Bare", placeholders: :none)
      violations = Tempfile.create(["bare", ".pptx"]) do |file|
        file.close
        deck.save(file.path)
        schema_violations(file.path)
      end
      expect(violations).to eq({})
    end

    # Guards the schema oracle itself: if this did not fail, a violation
    # anywhere else would not be detected either.
    it "detects a part that violates the schema" do
      master = deck.slide_masters.add(name: "Corporate")
      # p:clrMap is required and its attributes are enumerated.
      master.element.clrMap.set("bg1", "not-a-theme-colour")
      violations = Tempfile.create(["broken", ".pptx"]) do |file|
        file.close
        deck.save(file.path)
        schema_violations(file.path)
      end
      expect(violations.keys).to include("ppt/slideMasters/slideMaster2.xml")
    end
  end
end
