# frozen_string_literal: true

require "open3"
require "ruby_pptx/refinements"

RSpec.describe "Ruby idioms" do
  let(:deck) { Pptx::Presentation.new_default }
  let(:slide) { deck.slides.add(deck.slide_layouts["Title and Content"]) }

  describe Pptx::Sliceable do
    before { 3.times { deck.slides.add(deck.slide_layouts["Blank"]) } }

    # Regression: `[]` used to hand the index straight to the backing list and
    # then treat the result as one element, so a range raised NoMethodError on
    # Array deep inside the library.
    it "takes a range as Array does" do
      slice = deck.slides[0..1]
      aggregate_failures do
        expect(slice).to be_an(Array)
        expect(slice.size).to eq(2)
        expect(slice).to all(be_a(Pptx::Slide))
        expect(slice.map(&:slide_id)).to eq(deck.slides.to_a.first(2).map(&:slide_id))
      end
    end

    it "takes a start and a length" do
      expect(deck.slides[1, 2].map(&:slide_id)).to eq(deck.slides.to_a[1, 2].map(&:slide_id))
    end

    it "still returns a single member for an integer index" do
      expect(deck.slides[1]).to be_a(Pptx::Slide)
    end

    it "keeps counting from the end" do
      expect(deck.slides[-1].slide_id).to eq(deck.slides.to_a.last.slide_id)
    end

    # Array semantics: an index past the end is nil, and so is a range that
    # starts past it -- not an empty array.
    it "matches Array at the edges" do
      aggregate_failures do
        expect(deck.slides[99]).to be_nil
        expect(deck.slides[99..]).to be_nil
        expect(deck.slides[3..99].size).to eq(0)
      end
    end

    it "slices every collection, not just slides" do
      master = deck.slide_masters[0]
      table = slide.shapes.add_table(3, 4, at: [0, 0], size: [Pptx.inches(9), Pptx.inches(3)]).table
      aggregate_failures do
        expect(deck.slide_layouts[0..1]).to all(be_a(Pptx::SlideLayout))
        expect(deck.slide_masters[0..0]).to all(be_a(Pptx::SlideMaster))
        expect(master.placeholders[0..1]).to all(be_a(Pptx::MasterPlaceholder))
        expect(slide.shapes[0..0]).to all(be_a(Pptx::BaseShape))
        expect(table.rows[0..1].size).to eq(2)
        expect(table.columns[0..1].size).to eq(2)
      end
    end

    it "slices sections" do
      deck.sections.add("Intro")
      deck.sections.add("Body")
      expect(deck.sections[0..1].map(&:name)).to eq(%w[Intro Body])
    end

    # A String still means "by name" on the collections that support it.
    it "leaves name lookup alone" do
      deck.sections.add("Intro")
      aggregate_failures do
        expect(deck.slide_layouts["Blank"].name).to eq("Blank")
        expect(deck.sections["Intro"].name).to eq("Intro")
      end
    end
  end

  describe Pptx::PatternMatching do
    it "matches a shape on its type and name" do
      box = slide.shapes.add_shape(:rectangle, at: [0, 0], size: [Pptx.inches(6), Pptx.inches(1)])
      box.name = "Banner"
      matched = case box
                in { shape_type: :AUTO_SHAPE, name: } then name
                end
      expect(matched).to eq("Banner")
    end

    # An enum-valued attribute reads as its symbol in a pattern, which is the
    # whole point -- the accessor still hands back the member.
    it "reports an enum attribute as a symbol while the reader keeps the member" do
      box = slide.shapes.add_shape(:rectangle, at: [0, 0], size: [100, 100])
      aggregate_failures do
        expect(box.deconstruct_keys([:shape_type])).to eq(shape_type: :AUTO_SHAPE)
        expect(box.shape_type).to be_a(Pptx::Enum::Member)
      end
    end

    it "matches a placeholder through its nested format" do
      matched = case slide.shapes.title
                in { shape_type: :PLACEHOLDER, placeholder_format: { type: :CENTER_TITLE | :TITLE } }
                  :title
                end
      expect(matched).to eq(:title)
    end

    it "supports a guard on a Length" do
      box = slide.shapes.add_shape(:rectangle, at: [0, 0], size: [Pptx.inches(6), Pptx.inches(1)])
      matched = case box
                in { width: } if width > Pptx.inches(5) then :wide
                in { width: } then :narrow
                end
      expect(matched).to eq(:wide)
    end

    # Computing every key on every match would mean walking the text of a
    # shape to answer a question about its name.
    it "computes only the keys a pattern asks for" do
      box = slide.shapes.add_shape(:rectangle, at: [0, 0], size: [100, 100])
      aggregate_failures do
        expect(box.deconstruct_keys([:name]).keys).to eq([:name])
        expect(box.deconstruct_keys(nil).keys).to include(:shape_id, :name, :shape_type, :width)
      end
    end

    it "ignores a key the object does not offer" do
      box = slide.shapes.add_shape(:rectangle, at: [0, 0], size: [100, 100])
      expect(box.deconstruct_keys(%i[name nonsense])).to eq(name: box.name)
    end

    it "matches a slide, a layout and a presentation" do
      aggregate_failures do
        expect(case slide
               in { slide_id: Integer => id } then id
               end).to eq(slide.slide_id)
        expect(case deck.slide_layouts[0]
               in { name: String => n } then n
               end).to eq(deck.slide_layouts[0].name)
        expect(case deck
               in { slide_width:, slide_height: } then [slide_width, slide_height]
               end).to eq([deck.slide_width, deck.slide_height])
      end
    end

    it "matches text objects" do
      slide.shapes.title.text = "Hello"
      frame = slide.shapes.title.text_frame
      expect(case frame
             in { text: "Hello", paragraphs: [_] } then :ok
             end).to eq(:ok)
    end

    it "matches a chart" do
      data = Pptx::ChartData.new
      data.categories = %w[East West]
      data.add_series("Q1", [1, 2])
      frame = slide.shapes.add_chart(:column_clustered, data, at: [0, 0],
                                                              size: [Pptx.inches(6), Pptx.inches(4)])
      expect(case frame.chart
             in { plot_type: :BAR, categories: ["East", "West"] } then :ok
             end).to eq(:ok)
    end

    it "matches a Length in whatever unit reads best" do
      aggregate_failures do
        expect(case Pptx.inches(2)
               in { inches: } then inches
               end).to eq(2.0)
        expect(case Pptx.pt(18)
               in { pt: 18.0 } then :ok
               end).to eq(:ok)
      end
    end

    it "matches an RGBColor as keys or as an array" do
      color = Pptx::RGBColor["1F497D"]
      aggregate_failures do
        expect(case color
               in { r:, g:, b: } then [r, g, b]
               end).to eq([31, 73, 125])
        expect(case color
               in [r, _, _] then r
               end).to eq(31)
      end
    end

    it "deconstructs a collection into an array pattern" do
      slide
      deck.slides.add(deck.slide_layouts["Blank"])
      expect(case deck.slides
             in [first, *rest] then [first.class, rest.size]
             end).to eq([Pptx::Slide, 1])
    end
  end

  describe Pptx::Point do
    it "stands in for the array `at:` has always taken" do
      box = slide.shapes.add_shape(:rectangle,
                                   at: Pptx.point(Pptx.inches(1), Pptx.inches(2)),
                                   size: Pptx.size(Pptx.inches(3), Pptx.inches(1)))
      aggregate_failures do
        expect(box.left).to eq(Pptx.inches(1))
        expect(box.top).to eq(Pptx.inches(2))
        expect(box.width).to eq(Pptx.inches(3))
        expect(box.height).to eq(Pptx.inches(1))
      end
    end

    it "destructures like the array it replaces" do
      x, y = Pptx.point(Pptx.inches(1), Pptx.inches(2))
      expect([x, y]).to eq([Pptx.inches(1), Pptx.inches(2)])
    end

    # Length defines to_int, so `100 == Pptx.emu(100)` is already true --
    # only the class shows whether the value was actually normalised.
    it "normalises a bare number to a Length of EMU, as the array form does" do
      point = Pptx.point(100, 200)
      aggregate_failures do
        expect(point.x).to be_a(Pptx::Length)
        expect(point.x.emu).to eq(100)
        expect(point.y.emu).to eq(200)
      end
    end

    it "adds and subtracts, taking a Point or a plain array" do
      origin = Pptx.point(Pptx.inches(1), Pptx.inches(1))
      aggregate_failures do
        expect((origin + Pptx.point(Pptx.inches(1), 0)).x).to eq(Pptx.inches(2))
        expect((origin + [Pptx.inches(2), 0]).x).to eq(Pptx.inches(3))
        expect((origin - [Pptx.inches(1), 0]).x).to eq(Pptx.emu(0))
      end
    end

    it "is a value object: equal by contents and frozen" do
      aggregate_failures do
        expect(Pptx.point(1, 2)).to eq(Pptx.point(1, 2))
        expect(Pptx.point(1, 2)).to be_frozen
        expect({ Pptx.point(1, 2) => :here }[Pptx.point(1, 2)]).to eq(:here)
      end
    end

    it "builds from either a Point or an array" do
      expect(Pptx::Point.from([1, 2])).to eq(Pptx::Point.from(Pptx.point(1, 2)))
    end

    # Handing back the object itself, rather than an equal copy, is the point
    # of accepting a Point at all.
    it "returns the same object when it is already a Point" do
      point = Pptx.point(1, 2)
      aggregate_failures do
        expect(Pptx::Point.from(point)).to be(point)
        expect(Pptx::Size.from(Pptx.size(1, 2))).to be_a(Pptx::Size)
      end
    end

    it "keeps an existing Length rather than rebuilding it" do
      length = Pptx.inches(1)
      expect(Pptx.point(length, 0).x).to be(length)
    end

    # Length.coerce is strict, so a Point is too: a Float is not whole EMU.
    it "refuses a nil or fractional coordinate" do
      aggregate_failures do
        expect { Pptx.point(nil, 0) }.to raise_error(ArgumentError, /x is required/)
        expect { Pptx.point(1.5, 0) }.to raise_error(TypeError)
      end
    end
  end

  describe Pptx::Size do
    it "scales and reports its aspect ratio" do
      size = Pptx.size(Pptx.inches(4), Pptx.inches(2))
      aggregate_failures do
        expect((size * 2).width).to eq(Pptx.inches(8))
        expect(size.aspect_ratio).to eq(2.0)
      end
    end

    it "destructures like an array" do
      width, height = Pptx.size(Pptx.inches(3), Pptx.inches(1))
      expect([width, height]).to eq([Pptx.inches(3), Pptx.inches(1)])
    end
  end

  # The refinement and the core extension share one module, so they cannot
  # drift apart; only the scope differs.
  describe Pptx::Lengths do
    using Pptx::Lengths

    it "gives Numeric the length methods inside this file" do
      aggregate_failures do
        expect(1.inch).to eq(Pptx.inches(1))
        expect(2.5.cm).to eq(Pptx.cm(2.5))
        expect(18.pt).to eq(Pptx.pt(18))
        expect(914_400.emu).to eq(Pptx.inches(1))
        expect(1.centipoints).to eq(Pptx.centipoints(1))
      end
    end

    it "spells out the plural and singular forms" do
      aggregate_failures do
        expect(72.points).to eq(1.inch)
        expect(72.point).to eq(1.inch)
        expect(1.inches).to eq(1.inch)
      end
    end

    it "reaches the shape API" do
      box = slide.shapes.add_shape(:rectangle, at: [1.inch, 2.inch], size: [3.inch, 1.inch])
      expect(box.width).to eq(Pptx.inches(3))
    end
  end

  # In a subprocess, because another spec file requires pptx/core_ext, which
  # patches Numeric for the whole process -- which is the hazard the
  # refinement exists to avoid, and would make this assertion order-dependent
  # if it were made in here.
  it "leaves Numeric alone outside a `using` scope" do
    script = <<~RUBY
      require "ruby_pptx"
      require "ruby_pptx/refinements"
      outside = 1.respond_to?(:inch)
      inside = Module.new do
        using Pptx::Lengths
        def self.check = 1.respond_to?(:inch)
      end.check
      print [outside, inside].inspect
    RUBY
    out, err, status = Open3.capture3("ruby", "-Ilib", "-e", script)
    raise "refinement probe failed: #{err}" unless status.success?

    expect(out).to eq("[false, true]")
  end

  describe "value objects defined with Data" do
    it "are immutable" do
      operation = Pptx::FreeformBuilder::Operation.new(:line, 1, 2)
      aggregate_failures do
        expect(operation).to be_frozen
        expect(operation.kind).to eq(:line)
        expect { operation.instance_variable_set(:@kind, :move) }.to raise_error(FrozenError)
      end
    end

    it "compare by value" do
      expect(Pptx::FreeformBuilder::Operation.new(:line, 1, 2))
        .to eq(Pptx::FreeformBuilder::Operation.new(:line, 1, 2))
    end
  end
end
