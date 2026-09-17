# frozen_string_literal: true

RSpec.describe Pptx::Oxml::ContentModel do
  # The test schema below claims tags the real element classes also claim, such
  # as "a:xfrm". Defining a class registers it immediately, and RSpec loads
  # every spec file before running any example, so simply restoring in
  # after(:all) would still leave the schema installed for whatever runs first.
  # Instead: snapshot, define, put the real table straight back, and swap the
  # schema in only for the duration of this file's examples.
  pristine_registry = Pptx::Oxml::Registry.registered


  # A cut-down but faithful slice of the real schema: <a:xfrm> holds an
  # optional <a:off> then an optional <a:ext>, and carries a `rot` attribute.
  # These register themselves globally, so the registry is restored afterwards.
  module Schema
    ST = Pptx::Oxml::SimpleTypes

    class CT_BodyPr < Pptx::Oxml::Element
      tag "a:bodyPr"
      optional_attr "anchor", type: Pptx::Enum::MSO_ANCHOR
    end

    class CT_Point2D < Pptx::Oxml::Element
      tag "a:off"
      required_attr "x", type: ST::ST_Coordinate
      required_attr "y", type: ST::ST_Coordinate
    end

    class CT_PositiveSize2D < Pptx::Oxml::Element
      tag "a:ext"
      required_attr "cx", type: ST::ST_PositiveCoordinate
      required_attr "cy", type: ST::ST_PositiveCoordinate
    end

    class CT_Transform2D < Pptx::Oxml::Element
      tag "a:xfrm"
      optional_attr "rot", type: ST::ST_Angle, default: 0.0
      optional_attr "flipH", type: ST::XsdBoolean, default: false
      zero_or_one "a:off", successors: %w[a:ext]
      zero_or_one "a:ext", successors: []
    end

    class CT_Blip < Pptx::Oxml::Element
      tag "a:blip"
      optional_attr "r:embed", type: ST::ST_RelationshipId
    end

    class CT_Paragraph < Pptx::Oxml::Element
      tag "a:p"
      zero_or_one "a:pPr", successors: %w[a:r a:endParaRPr]
      zero_or_more "a:r", successors: %w[a:endParaRPr]
      zero_or_one "a:endParaRPr", successors: []
    end

    class CT_FillProps < Pptx::Oxml::Element
      tag "a:spPr"
      zero_or_one_choice [choice("a:noFill"), choice("a:solidFill"), choice("a:gradFill")],
                         successors: %w[a:ln], as: :eg_fillProperties
      zero_or_one "a:ln", successors: []
    end
  end

  schema_registry = Pptx::Oxml::Registry.registered
  Pptx::Oxml::Registry.reset!(pristine_registry)

  before(:all) { Pptx::Oxml::Registry.reset!(schema_registry) }
  after(:all) { Pptx::Oxml::Registry.reset!(pristine_registry) }

  def element(xml) = Pptx::Oxml::Element.parse(xml)

  def xfrm(inner = "")
    element(%(<a:xfrm #{Pptx::Oxml::Ns.nsdecls('a', 'r')}>#{inner}</a:xfrm>))
  end

  describe "dispatch" do
    it "wraps a node in the class registered for its tag" do
      expect(xfrm).to be_a(Schema::CT_Transform2D)
    end

    it "wraps an unmodelled tag in the base Element" do
      expect(element(%(<a:unknown #{Pptx::Oxml::Ns.nsdecls('a')}/>))).to be_an_instance_of(Pptx::Oxml::Element)
    end

    it "returns equal wrappers for the same node" do
      parent = xfrm("<a:off x='1' y='2'/>")
      expect(parent.off).to eq(parent.off)
      expect(parent.off.hash).to eq(parent.off.hash)
    end
  end

  describe "optional_attr" do
    it "returns the default when absent" do
      expect(xfrm.rot).to eq(0.0)
      expect(xfrm.flipH).to be(false)
    end

    it "type-converts when present" do
      expect(xfrm.tap { |e| e.set("rot", "2700000") }.rot).to eq(45.0)
    end

    it "round-trips an assignment through the simple type" do
      e = xfrm
      e.rot = 45.0
      expect(e.get("rot")).to eq("2700000")
      expect(e.rot).to eq(45.0)
    end

    it "removes the attribute when assigned its default" do
      e = xfrm
      e.rot = 45.0
      e.rot = 0.0
      expect(e.get("rot")).to be_nil
      expect(e.xml).not_to include("rot")
    end

    it "handles a namespace-prefixed attribute" do
      blip = element(%(<a:blip #{Pptx::Oxml::Ns.nsdecls('a', 'r')}/>))
      blip.embed = "rId3"
      aggregate_failures do
        expect(blip.embed).to eq("rId3")
        expect(blip.xml).to include('r:embed="rId3"')
        expect(blip.node.attribute_with_ns("embed", Pptx::Oxml::Ns.nsuri("r"))).not_to be_nil
      end
    end
  end

  describe "required_attr" do
    it "type-converts when present" do
      expect(xfrm("<a:off x='914400' y='0'/>").off.x).to eq(Pptx.inches(1))
    end

    it "raises InvalidXmlError when missing" do
      expect { xfrm("<a:off y='0'/>").off.x }
        .to raise_error(Pptx::InvalidXmlError, /required "x" attribute not present/)
    end
  end

  describe "one_and_only_one" do
    it "raises when the required child is absent" do
      klass = Class.new(Pptx::Oxml::Element) { one_and_only_one "a:ext" }
      instance = klass.new(xfrm.node)
      expect { instance.ext }.to raise_error(Pptx::InvalidXmlError, /required <a:ext>/)
    end
  end

  describe "zero_or_one" do
    it "returns nil when absent and the element when present" do
      expect(xfrm.off).to be_nil
      expect(xfrm("<a:off x='0' y='0'/>").off).to be_a(Schema::CT_Point2D)
    end

    it "get_or_add is idempotent" do
      e = xfrm
      first = e.get_or_add_off
      expect(e.get_or_add_off).to eq(first)
      expect(e.node.element_children.size).to eq(1)
    end

    it "add accepts attributes as keywords" do
      e = xfrm
      e.add_off(x: Pptx.inches(1), y: Pptx.emu(0))
      expect(e.xml).to include('x="914400"', 'y="0"')
    end

    it "remove clears the child" do
      e = xfrm("<a:off x='0' y='0'/>")
      e.remove_off
      expect(e.off).to be_nil
    end

    it "inserts in schema order regardless of call order" do
      e = xfrm
      e.get_or_add_ext          # the later sibling, added first
      e.get_or_add_off
      expect(e.node.element_children.map(&:name)).to eq(%w[off ext])
    end
  end

  describe "zero_or_more" do
    it "exposes a list rather than a singular accessor" do
      p = element(%(<a:p #{Pptx::Oxml::Ns.nsdecls('a')}><a:r/><a:r/></a:p>))
      aggregate_failures do
        expect(p.r_list.size).to eq(2)
        expect(p).not_to respond_to(:r)
      end
    end

    it "keeps repeated children before their successor" do
      p = element(%(<a:p #{Pptx::Oxml::Ns.nsdecls('a')}/>))
      p.get_or_add_endParaRPr
      p.add_r
      p.add_r
      p.get_or_add_pPr
      expect(p.node.element_children.map(&:name)).to eq(%w[pPr r r endParaRPr])
    end
  end

  describe "zero_or_one_choice" do
    subject(:sp_pr) { element(%(<a:spPr #{Pptx::Oxml::Ns.nsdecls('a')}/>)) }

    it "reports the member present, or nil" do
      expect(sp_pr.eg_fillProperties).to be_nil
      sp_pr.add_solidFill
      expect(sp_pr.eg_fillProperties.nsptag).to eq("a:solidFill")
    end

    it "swaps one member for another, never leaving two" do
      sp_pr.get_or_change_to_solidFill
      sp_pr.get_or_change_to_gradFill
      aggregate_failures do
        expect(sp_pr.node.element_children.map(&:name)).to eq(%w[gradFill])
        expect(sp_pr.solidFill).to be_nil
      end
    end

    it "is idempotent when the member is already present" do
      first = sp_pr.get_or_change_to_noFill
      expect(sp_pr.get_or_change_to_noFill).to eq(first)
    end

    it "removes the whole group" do
      sp_pr.add_solidFill
      sp_pr.remove_eg_fillProperties
      expect(sp_pr.eg_fillProperties).to be_nil
    end

    it "inserts the chosen member before the group's successor" do
      sp_pr.get_or_add_ln
      sp_pr.get_or_change_to_solidFill
      expect(sp_pr.node.element_children.map(&:name)).to eq(%w[solidFill ln])
    end
  end

  # python-pptx uses its enumerations directly as attribute types; ours satisfy
  # the same from_xml/to_xml contract, so they drop straight into the DSL.
  describe "an enumeration used as an attribute type" do
    subject(:body_pr) { element(%(<a:bodyPr #{Pptx::Oxml::Ns.nsdecls('a')}/>)) }

    it "reads an XML value as the enum member" do
      body_pr.set("anchor", "ctr")
      expect(body_pr.anchor).to eq(Pptx::Enum::MSO_ANCHOR::MIDDLE)
    end

    it "writes the member's XML value" do
      body_pr.anchor = Pptx::Enum::MSO_ANCHOR::BOTTOM
      expect(body_pr.get("anchor")).to eq("b")
    end

    it "returns nil when the attribute is absent" do
      expect(body_pr.anchor).to be_nil
    end

    it "removes the attribute when assigned nil" do
      body_pr.anchor = Pptx::Enum::MSO_ANCHOR::TOP
      body_pr.anchor = nil
      expect(body_pr.get("anchor")).to be_nil
    end

    it "rejects a value outside the enumeration" do
      expect { body_pr.anchor = 99 }.to raise_error(ArgumentError, /not a member/)
    end
  end

  describe "document safety" do
    it "creates children in the parent's own document" do
      e = xfrm
      expect(e.build("a:off").node.document).to equal(e.node.document)
    end

    it "refuses an element built in a different document" do
      # Nokogiri re-creates a node adopted across documents, so the object
      # inserted is not the object built; better to refuse than to hand back a
      # wrapper that silently points at an orphan.
      foreign = xfrm.build("a:off")
      expect { xfrm.append(foreign) }
        .to raise_error(ArgumentError, /different document/)
    end

    it "does not redeclare a namespace already in scope" do
      e = xfrm
      e.get_or_add_off
      expect(e.xml.scan(/xmlns:a=/).size).to eq(1)
    end
  end
end
