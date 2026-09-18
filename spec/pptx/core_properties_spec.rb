# frozen_string_literal: true

require "json"
require "open3"

RSpec.describe Pptx::Oxml::CT_CoreProperties do
  def oracle_path = File.expand_path("../../tools/coreprops_oracle.py", __dir__)

  subject(:core_props) { described_class.new_element }

  def python_core_props(ops)
    out, err, status = Open3.capture3("python3", oracle_path, stdin_data: JSON.dump(ops: ops))
    raise "core-props oracle failed (#{status.exitstatus}): #{err}" if out.empty?

    out
  end

  describe "text properties" do
    it "reads an empty string when the element is absent" do
      expect(core_props.title).to eq("")
    end

    it "round-trips a value" do
      core_props.title = "Quarterly Review"
      expect(core_props.title).to eq("Quarterly Review")
    end

    it "replaces rather than appends on a second assignment" do
      core_props.author = "First"
      core_props.author = "Second"
      aggregate_failures do
        expect(core_props.author).to eq("Second")
        expect(core_props.xml.scan("<dc:creator>").size).to eq(1)
      end
    end

    it "refuses a value over the 255-character schema limit" do
      expect { core_props.title = "x" * 256 }
        .to raise_error(ArgumentError, /255-character limit/)
    end

    it "accepts exactly 255 characters" do
      expect { core_props.title = "x" * 255 }.not_to raise_error
    end
  end

  describe "#revision" do
    it "is 0 when absent" do
      expect(core_props.revision).to eq(0)
    end

    it "round-trips a positive integer" do
      core_props.revision = 42
      expect(core_props.revision).to eq(42)
    end

    it "reads a non-integer or negative value as 0" do
      aggregate_failures do
        core_props.get_or_add_revision.text = "not a number"
        expect(core_props.revision).to eq(0)
        core_props.get_or_add_revision.text = "-5"
        expect(core_props.revision).to eq(0)
      end
    end

    it "refuses a non-positive assignment" do
      aggregate_failures do
        expect { core_props.revision = 0 }.to raise_error(ArgumentError, /positive Integer/)
        expect { core_props.revision = "3" }.to raise_error(ArgumentError, /positive Integer/)
      end
    end
  end

  describe "datetime properties" do
    it "is nil when absent" do
      expect(core_props.modified).to be_nil
    end

    it "round-trips a Time as UTC" do
      core_props.modified = Time.utc(2026, 9, 17, 12, 0, 0)
      expect(core_props.modified).to eq(Time.utc(2026, 9, 17, 12, 0, 0))
    end

    it "converts a zoned Time to UTC on the way in" do
      core_props.modified = Time.new(2026, 9, 17, 14, 0, 0, "+02:00")
      expect(core_props.modified).to eq(Time.utc(2026, 9, 17, 12, 0, 0))
    end

    it "refuses anything that is not a Time" do
      expect { core_props.modified = "2026-09-17" }.to raise_error(TypeError, /Time/)
    end

    it "reads an unparseable stamp as nil rather than raising" do
      core_props.get_or_add_modified.text = "not a date"
      expect(core_props.modified).to be_nil
    end

    # dcterms:created and dcterms:modified must carry xsi:type; the others
    # must not. PowerPoint declares xsi on the root element, not on the child
    # that uses it, and C14N renders a declaration where it is declared.
    it "declares xsi on the root and types only the dcterms elements" do
      core_props.created = Time.utc(2026, 1, 2, 3, 4, 5)
      core_props.last_printed = Time.utc(2026, 1, 2, 3, 4, 5)
      xml = core_props.xml

      aggregate_failures do
        expect(xml.lines.first).to include('xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"')
        expect(xml).to match(/<dcterms:created xsi:type="dcterms:W3CDTF">/)
        expect(xml).to match(/<cp:lastPrinted>2026/)
        expect(xml.scan("xmlns:xsi=").size).to eq(1)
      end
    end
  end

  describe "W3CDTF parsing" do
    {
      "2003-12-31T10:14:55Z" => [2003, 12, 31, 10, 14, 55],
      "2003-12-31T10:14:55-08:00" => [2003, 12, 31, 18, 14, 55],
      "2003-12-31T10:14:55+02:00" => [2003, 12, 31, 8, 14, 55],
      "2003-12-31" => [2003, 12, 31, 0, 0, 0],
      "2003-12" => [2003, 12, 1, 0, 0, 0],
      "2003" => [2003, 1, 1, 0, 0, 0]
    }.each do |value, expected|
      it "parses #{value.inspect}" do
        expect(described_class.parse_w3cdtf(value)).to eq(Time.utc(*expected))
      end
    end

    # Ruby's strptime parses a prefix where Python's requires a full match, so
    # without a leftover check "%Y" would swallow a full timestamp.
    it "does not let a short format swallow a full timestamp" do
      expect(described_class.parse_w3cdtf("2003-12-31T10:14:55"))
        .to eq(Time.utc(2003, 12, 31, 10, 14, 55))
    end

    it "returns nil for an unparseable value" do
      expect(described_class.parse_w3cdtf("last Tuesday")).to be_nil
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    # Exercises a core-properties part built from scratch, where xsi is not
    # already declared -- the case a package that already has core properties
    # never reaches.
    it "serializes a newly built part exactly as python-pptx does" do
      ops = [
        ["title", "Ported to Ruby"],
        ["author", "Andi"],
        ["keywords", "ooxml, ruby"],
        ["revision", 7],
        ["created", [2026, 1, 2, 3, 4, 5]],
        ["modified", [2026, 9, 17, 12, 0, 0]],
        ["last_printed", [2025, 6, 1, 0, 0, 0]]
      ]

      core_props.title = "Ported to Ruby"
      core_props.author = "Andi"
      core_props.keywords = "ooxml, ruby"
      core_props.revision = 7
      core_props.created = Time.utc(2026, 1, 2, 3, 4, 5)
      core_props.modified = Time.utc(2026, 9, 17, 12, 0, 0)
      core_props.last_printed = Time.utc(2025, 6, 1, 0, 0, 0)

      expect(canonical(Pptx::Opc::Oxml.serialize_part_xml(core_props)))
        .to eq(canonical(python_core_props(ops)))
    end
  end

  def canonical(xml)
    Nokogiri::XML(xml, &:noblanks).canonicalize(Nokogiri::XML::XML_C14N_1_0)
  end
end
