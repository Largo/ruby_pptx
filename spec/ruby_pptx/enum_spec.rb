# frozen_string_literal: true

require "json"
require "open3"

RSpec.describe Pptx::Enum do
  def oracle_path = File.expand_path("../../tools/enum_oracle.py", __dir__)

  def python_enums
    out, err, status = Open3.capture3("python3", oracle_path)
    raise "enum oracle failed (#{status.exitstatus}): #{err}" if out.empty?

    JSON.parse(out)
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    let(:expected) { python_enums }

    it "defines every enumeration, including the MS API aliases" do
      missing = expected.keys.reject { |name| Pptx::Enum.const_defined?(name, false) }
      expect(missing).to be_empty
    end

    it "gives every member the same name, MS API value and XML value" do
      mismatches = expected.flat_map do |enum_name, spec|
        next ["#{enum_name}: not defined"] unless Pptx::Enum.const_defined?(enum_name, false)

        compare_members(Pptx::Enum.const_get(enum_name), enum_name, spec["members"])
      end

      expect(mismatches).to be_empty, -> { mismatches.first(20).join("\n") }
    end

    it "points each alias at the very same class as python-pptx does" do
      mismatches = expected.filter_map do |enum_name, spec|
        next if enum_name == spec["canonical"]

        ours = Pptx::Enum.const_get(enum_name)
        canonical = Pptx::Enum.const_get(spec["canonical"])
        "#{enum_name} should alias #{spec["canonical"]}" unless ours.equal?(canonical)
      end

      expect(mismatches).to be_empty
    end

    it "writes the same XML value for every XML-mapped member" do
      failures = expected.flat_map do |enum_name, spec|
        next [] unless spec["xml_mapped"]

        enum = Pptx::Enum.const_get(enum_name)
        spec["members"].filter_map do |m|
          next if m["xml"].nil?

          actual = enum.to_xml(enum[m["name"].to_sym])
          unless actual == m["xml"]
            "#{enum_name}.#{m["name"]}.to_xml => #{actual.inspect}, want #{m["xml"].inspect}"
          end
        end
      end

      expect(failures).to be_empty, -> { failures.first(20).join("\n") }
    end

    # Some XML values map to more than one MS API member, so `from_xml` is
    # inherently lossy; what matters is that we resolve to the same member
    # python-pptx does.
    it "resolves every XML value to the same member python-pptx does" do
      failures = expected.flat_map do |enum_name, spec|
        next [] unless spec["xml_mapped"]

        enum = Pptx::Enum.const_get(enum_name)
        spec["from_xml"].filter_map do |xml_value, member_name|
          actual = enum.from_xml(xml_value).name.to_s
          unless actual == member_name
            "#{enum_name}.from_xml(#{xml_value.inspect}) => #{actual}, want #{member_name}"
          end
        end
      end

      expect(failures).to be_empty, -> { failures.first(20).join("\n") }
    end

    it "covers every member python-pptx defines" do
      total = expected.values.sum { |spec| spec["members"].size }
      ours = expected.keys.sum { |name| Pptx::Enum.const_get(name).count }
      expect(ours).to eq(total)
    end
  end

  def compare_members(enum, enum_name, members)
    members.filter_map do |m|
      ours = enum[m["name"].to_sym]
      next "#{enum_name}.#{m["name"]}: missing" if ours.nil?
      next "#{enum_name}.#{m["name"]}: value #{ours.value} != #{m["value"]}" if ours.value != m["value"]

      expected_xml = m["xml"]
      actual_xml = ours.xml_value? ? ours.xml_value : nil
      next if actual_xml == expected_xml

      "#{enum_name}.#{m["name"]}: xml #{actual_xml.inspect} != #{expected_xml.inspect}"
    end
  end

  describe "Ruby-side behaviour" do
    it "compares equal to its MS API integer" do
      expect(Pptx::Enum::MSO_ANCHOR::MIDDLE).to eq(3)
    end

    it "is Enumerable" do
      expect(Pptx::Enum::MSO_ANCHOR.map(&:name)).to include(:TOP, :MIDDLE, :BOTTOM)
    end

    it "looks members up by name, integer or member" do
      middle = Pptx::Enum::MSO_ANCHOR::MIDDLE
      aggregate_failures do
        expect(Pptx::Enum::MSO_ANCHOR[:MIDDLE]).to eq(middle)
        expect(Pptx::Enum::MSO_ANCHOR[3]).to eq(middle)
        expect(Pptx::Enum::MSO_ANCHOR[middle]).to eq(middle)
        expect(Pptx::Enum::MSO_ANCHOR[:NOPE]).to be_nil
      end
    end

    it "refuses to write a return-value-only member to XML" do
      # MIXED exists to be returned, never written; it has no XML value.
      expect { Pptx::Enum::MSO_THEME_COLOR.to_xml(Pptx::Enum::MSO_THEME_COLOR::MIXED) }
        .to raise_error(ArgumentError, /no XML representation/)
    end

    it "resolves a shared XML value to the first member declared" do
      # 'borderCallout3' is the XML for both LINE_CALLOUT_3 and LINE_CALLOUT_4.
      expect(Pptx::Enum::MSO_SHAPE.from_xml("borderCallout3").name).to eq(:LINE_CALLOUT_3)
      expect(Pptx::Enum::MSO_SHAPE.to_xml(Pptx::Enum::MSO_SHAPE::LINE_CALLOUT_4)).to eq("borderCallout3")
    end

    it "raises on an unmapped XML value" do
      expect { Pptx::Enum::MSO_ANCHOR.from_xml("sideways") }
        .to raise_error(ArgumentError, /no XML mapping/)
      expect { Pptx::Enum::MSO_ANCHOR.from_xml("") }
        .to raise_error(ArgumentError, /no XML mapping/)
    end

    it "renders as name and value" do
      expect(Pptx::Enum::MSO_ANCHOR::MIDDLE.to_s).to eq("MIDDLE (3)")
    end
  end
end
