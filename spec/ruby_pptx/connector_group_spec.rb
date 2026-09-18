# frozen_string_literal: true

RSpec.describe Pptx::Connector do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }

  def box(left, top)
    slide.shapes.add_shape(:rounded_rectangle, at: [Pptx.inches(left), Pptx.inches(top)],
                                               size: [Pptx.inches(2), Pptx.inches(1)])
  end

  def connector(begin_at, end_at, type: :straight)
    slide.shapes.add_connector(type, begin_at: begin_at.map { |v| Pptx.inches(v) },
                                     end_at: end_at.map { |v| Pptx.inches(v) })
  end

  describe "adding one" do
    subject(:line) { connector([1, 1], [4, 3]) }

    it "is a connector shape with the given end points" do
      aggregate_failures do
        expect(line).to be_a(described_class)
        expect(line.shape_type).to eq(Pptx::Enum::MSO_SHAPE_TYPE::LINE)
        expect(line.name).to eq("Connector 1")
        expect(line.begin_x).to eq(Pptx.inches(1))
        expect(line.end_y).to eq(Pptx.inches(3))
      end
    end

    it "accepts a connector type by symbol" do
      expect { connector([0, 0], [1, 1], type: :elbow) }.not_to raise_error
    end

    # A connector is a bounding box plus flip flags, not two points, so a line
    # running right to left is the same box with flipH set rather than a
    # negative width.
    it "stores a right-to-left line as a flipped box" do
      reversed = connector([6, 1], [2, 1])
      aggregate_failures do
        expect(reversed.begin_x).to eq(Pptx.inches(6))
        expect(reversed.end_x).to eq(Pptx.inches(2))
        expect(reversed.width).to eq(Pptx.inches(4))
        expect(reversed.element.flipH).to be(true)
      end
    end

    it "stores a bottom-to-top line as a vertically flipped box" do
      reversed = connector([1, 5], [1, 2])
      aggregate_failures do
        expect(reversed.begin_y).to eq(Pptx.inches(5))
        expect(reversed.end_y).to eq(Pptx.inches(2))
        expect(reversed.element.flipV).to be(true)
      end
    end
  end

  describe "moving an end point" do
    subject(:line) { connector([1, 1], [4, 1]) }

    it "keeps the other end where it was" do
      line.begin_x = Pptx.inches(2)
      aggregate_failures do
        expect(line.begin_x).to eq(Pptx.inches(2))
        expect(line.end_x).to eq(Pptx.inches(4))
      end
    end

    # Dragging an end past the other one has to flip the box, since a box
    # cannot have a negative width.
    it "flips the box when an end is dragged past the other" do
      line.begin_x = Pptx.inches(6)
      aggregate_failures do
        expect(line.begin_x).to eq(Pptx.inches(6))
        expect(line.end_x).to eq(Pptx.inches(4))
        expect(line.element.flipH).to be(true)
        expect(line.width).to eq(Pptx.inches(2))
      end
    end

    it "round-trips every end point independently" do
      line.begin_x = Pptx.inches(2)
      line.begin_y = Pptx.inches(3)
      line.end_x = Pptx.inches(7)
      line.end_y = Pptx.inches(5)
      aggregate_failures do
        expect([line.begin_x, line.begin_y]).to eq([Pptx.inches(2), Pptx.inches(3)])
        expect([line.end_x, line.end_y]).to eq([Pptx.inches(7), Pptx.inches(5)])
      end
    end
  end

  describe "connecting to shapes" do
    subject(:line) { connector([0, 0], [1, 1]) }

    it "records the attachment and moves the end to the connection point" do
      target = box(2, 2)
      line.begin_connect(target, 0)
      aggregate_failures do
        expect(line).to be_begin_connected
        # Point 0 is the top centre of the bounding box.
        expect(line.begin_x).to eq(Pptx.inches(3))
        expect(line.begin_y).to eq(Pptx.inches(2))
      end
    end

    it "puts each of the four points where PowerPoint does" do
      target = box(2, 2) # 2in wide, 1in tall, so centre is (3, 2.5)
      {
        0 => [3.0, 2.0],   # top centre
        1 => [2.0, 2.5],   # left middle
        2 => [3.0, 3.0],   # bottom centre
        3 => [4.0, 2.5]    # right middle
      }.each do |index, (x, y)|
        line.end_connect(target, index)
        expect([line.end_x.inches, line.end_y.inches]).to eq([x, y]),
                                                          "point #{index}"
      end
    end

    it "connects both ends to different shapes" do
      line.begin_connect(box(1, 1), 3)
      line.end_connect(box(5, 3), 1)
      aggregate_failures do
        expect(line).to be_begin_connected
        expect(line).to be_end_connected
        expect(line.element.xml).to include("<a:stCxn", "<a:endCxn")
      end
    end

    it "rejects a connection point that does not exist" do
      expect { line.begin_connect(box(1, 1), 9) }
        .to raise_error(ArgumentError, /out of range/)
    end
  end
end

RSpec.describe Pptx::GroupShape do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }

  def box(left, top, width = 2, height = 1)
    slide.shapes.add_shape(:rounded_rectangle, at: [Pptx.inches(left), Pptx.inches(top)],
                                               size: [Pptx.inches(width), Pptx.inches(height)])
  end

  it "adds an empty group" do
    group = slide.shapes.add_group_shape
    aggregate_failures do
      expect(group).to be_a(described_class)
      expect(group.shape_type).to eq(Pptx::Enum::MSO_SHAPE_TYPE::GROUP)
      expect(group.shapes.size).to eq(0)
    end
  end

  # The shapes move into the group, so the slide is left holding the group
  # alone rather than the group plus its members.
  it "takes the given shapes out of the slide and into the group" do
    first = box(1, 1)
    second = box(5, 3)
    group = slide.shapes.add_group_shape([first, second])
    aggregate_failures do
      expect(group.shapes.size).to eq(2)
      expect(slide.shapes.size).to eq(1)
      expect(slide.shapes.first).to eq(group)
    end
  end

  # A group has no position of its own; it is the bounding box of what it
  # holds, which is why it has to be recomputed on every change.
  it "sizes itself around its contents" do
    group = slide.shapes.add_group_shape([box(1, 1), box(5, 3)])
    aggregate_failures do
      expect(group.left).to eq(Pptx.inches(1))
      expect(group.top).to eq(Pptx.inches(1))
      expect(group.width).to eq(Pptx.inches(6))  # 1 to 7
      expect(group.height).to eq(Pptx.inches(3)) # 1 to 4
    end
  end

  it "grows when a shape is added inside it" do
    group = slide.shapes.add_group_shape([box(2, 2)])
    expect do
      group.shapes.add_shape(:oval, at: [0, 0], size: [Pptx.inches(1), Pptx.inches(1)])
    end.to change { group.left }.from(Pptx.inches(2)).to(Pptx.emu(0))
  end

  it "keeps its child coordinate space in step with its extents" do
    group = slide.shapes.add_group_shape([box(1, 1), box(5, 3)])
    xfrm = group.element.xfrm
    aggregate_failures do
      expect(xfrm.chOff.x).to eq(group.left)
      expect(xfrm.chExt.cx).to eq(group.width)
    end
  end

  it "resizes an enclosing group when a nested group changes" do
    outer = slide.shapes.add_group_shape([box(4, 4)])
    inner = outer.shapes.add_group_shape
    inner.shapes.add_shape(:oval, at: [Pptx.inches(1), Pptx.inches(1)],
                                  size: [Pptx.inches(1), Pptx.inches(1)])
    expect(outer.left).to eq(Pptx.inches(1))
  end

  # PowerPoint offers neither on a group; the shapes inside carry their own.
  it "has no text frame and no click action" do
    group = slide.shapes.add_group_shape
    aggregate_failures do
      expect(group.text_frame?).to be(false)
      expect { group.click_action }.to raise_error(Pptx::Error, /cannot have a click action/)
      expect { group.hyperlink }.to raise_error(Pptx::Error, /cannot have a hyperlink/)
    end
  end
end

RSpec.describe "connectors and groups agreement with python-pptx" do
  before { require_oracle! }

  it "adds connectors identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Inches
        from pptx.enum.shapes import MSO_CONNECTOR_TYPE, MSO_SHAPE
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        a = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(1), Inches(1), Inches(2), Inches(1))
        b = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(5), Inches(3), Inches(2), Inches(1))
        c = slide.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT,
                                       Inches(1), Inches(1), Inches(4), Inches(3))
        c.begin_connect(a, 3)
        c.end_connect(b, 1)
        slide.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT,
                                   Inches(6), Inches(1), Inches(2), Inches(1))
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        a = slide.shapes.add_shape(:rounded_rectangle, at: [Pptx.inches(1), Pptx.inches(1)],
                                                       size: [Pptx.inches(2), Pptx.inches(1)])
        b = slide.shapes.add_shape(:rounded_rectangle, at: [Pptx.inches(5), Pptx.inches(3)],
                                                       size: [Pptx.inches(2), Pptx.inches(1)])
        c = slide.shapes.add_connector(:straight, begin_at: [Pptx.inches(1), Pptx.inches(1)],
                                                  end_at: [Pptx.inches(4), Pptx.inches(3)])
        c.begin_connect(a, 3)
        c.end_connect(b, 1)
        slide.shapes.add_connector(:straight, begin_at: [Pptx.inches(6), Pptx.inches(1)],
                                              end_at: [Pptx.inches(2), Pptx.inches(1)])
        prs.save(path)
      }
    )
  end

  it "adds groups identically to python-pptx" do
    expect_same_package(
      python: <<~PY,
        from pptx.util import Inches
        from pptx.enum.shapes import MSO_SHAPE
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        a = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(1), Inches(1), Inches(2), Inches(1))
        b = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                   Inches(5), Inches(3), Inches(2), Inches(1))
        group = slide.shapes.add_group_shape([a, b])
        group.shapes.add_shape(MSO_SHAPE.CHEVRON,
                               Inches(2), Inches(2), Inches(1), Inches(1))
        prs.save(out)
      PY
      ruby: lambda { |path|
        prs = Pptx::Presentation.new_default
        slide = prs.slides.add(prs.slide_layouts[6])
        a = slide.shapes.add_shape(:rounded_rectangle, at: [Pptx.inches(1), Pptx.inches(1)],
                                                       size: [Pptx.inches(2), Pptx.inches(1)])
        b = slide.shapes.add_shape(:oval, at: [Pptx.inches(5), Pptx.inches(3)],
                                          size: [Pptx.inches(2), Pptx.inches(1)])
        group = slide.shapes.add_group_shape([a, b])
        group.shapes.add_shape(:chevron, at: [Pptx.inches(2), Pptx.inches(2)],
                                         size: [Pptx.inches(1), Pptx.inches(1)])
        prs.save(path)
      }
    )
  end
end
