# frozen_string_literal: true

RSpec.describe Pptx::Oxml::SimpleTypes do
  simple_types = Pptx::Oxml::SimpleTypes
  define_method(:st) { Pptx::Oxml::SimpleTypes }

  # Every case is checked against python-pptx itself, so these assert real
  # upstream behaviour rather than a transcription of it.
  cases = [
    # angles are 60000ths of a degree, exposed as float degrees
    ["ST_Angle", :from_xml, "2700000"],
    ["ST_Angle", :to_xml, 45.0],
    ["ST_Angle", :to_xml, -45.0],
    ["ST_Angle", :to_xml, 405.0],
    ["ST_PositiveFixedAngle", :to_xml, -427.42],
    ["ST_PositiveFixedAngle", :to_xml, 42.42],

    # percentages: 1000ths of a percent, or a float with '%'
    ["ST_Percentage", :from_xml, "42000"],
    ["ST_Percentage", :from_xml, "42.0%"],
    ["ST_Percentage", :from_xml, "-5000"],
    ["ST_Percentage", :to_xml, 0.42],
    ["ST_Percentage", :to_xml, -0.05],
    ["ST_PositiveFixedPercentage", :to_xml, 1.0],

    # lengths
    ["ST_Coordinate", :from_xml, "914400"],
    ["ST_Coordinate", :from_xml, "1.5in"],
    ["ST_Coordinate", :from_xml, "-914400"],
    ["ST_Coordinate", :to_xml, 914_400],
    ["ST_Coordinate32", :from_xml, "2540"],
    ["ST_PositiveCoordinate", :from_xml, "914400"],
    ["ST_UniversalMeasure", :from_xml, "1.5in"],
    ["ST_UniversalMeasure", :from_xml, "25.4mm"],
    ["ST_UniversalMeasure", :from_xml, "72pt"],
    ["ST_UniversalMeasure", :from_xml, "1pc"],
    ["ST_LineWidth", :from_xml, "12700"],

    # text
    ["ST_TextSpacingPoint", :from_xml, "1800"],
    ["ST_TextSpacingPoint", :to_xml, 228_600],
    ["ST_TextSpacingPercentOrPercentString", :from_xml, "175000"],
    ["ST_TextSpacingPercentOrPercentString", :from_xml, "175%"],
    ["ST_TextSpacingPercentOrPercentString", :to_xml, 1.75],
    ["ST_TextFontScalePercentOrPercentString", :from_xml, "92500"],
    ["ST_TextFontScalePercentOrPercentString", :from_xml, "92.5%"],
    ["ST_TextFontScalePercentOrPercentString", :to_xml, 92.5],

    # percent-literal integer types
    ["ST_BubbleScale", :from_xml, "150%"],
    ["ST_BubbleScale", :from_xml, "150"],
    ["ST_GapAmount", :from_xml, "150%"],
    ["ST_Overlap", :from_xml, "-27%"],
    ["ST_LblOffset", :from_xml, "100%"],

    # strings, booleans, enumerations
    ["XsdBoolean", :from_xml, "1"],
    ["XsdBoolean", :from_xml, "true"],
    ["XsdBoolean", :from_xml, "0"],
    ["XsdBoolean", :from_xml, "false"],
    ["XsdBoolean", :to_xml, true],
    ["XsdBoolean", :to_xml, false],
    ["ST_HexColorRGB", :to_xml, "3f2a1b"],
    ["XsdInt", :to_xml, -42],
    ["XsdUnsignedByte", :to_xml, 255],

    # out-of-range and malformed input must fail on both sides
    ["XsdUnsignedByte", :to_xml, 256],
    ["ST_SlideId", :to_xml, 255],
    ["ST_TextFontSize", :to_xml, 99],
    ["ST_MarkerSize", :to_xml, 1],
    ["ST_Style", :to_xml, 49],
    ["ST_HexColorRGB", :to_xml, "zzzzzz"],
    ["ST_HexColorRGB", :to_xml, "3f2a1"],
    ["XsdBoolean", :from_xml, "yes"],
    ["ST_AxisUnit", :to_xml, 0.0],
    ["ST_LineWidth", :to_xml, 20_116_801],
    ["ST_SlideSizeCoordinate", :to_xml, 100]
  ].freeze
  define_method(:cases) { cases }

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "converts every case identically" do
      expected = python_simple_types(
        cases.map { |type, op, value| { type: type, op: op, value: value } }
      )

      mismatches = cases.each_with_index.filter_map do |(type, op, value), i|
        ours = evaluate(type, op, value)
        theirs = normalize_python(expected[i])
        next if comparable?(ours, theirs)

        "#{type}.#{op}(#{value.inspect}): python-pptx => #{theirs.inspect}, ours => #{ours.inspect}"
      end

      expect(mismatches).to be_empty, -> { mismatches.join("\n") }
    end
  end

  # Invoke our simple type, reducing a raised error to a marker so a case that
  # must fail on both sides can be compared like any other.
  def evaluate(type_name, op, value)
    result = st.const_get(type_name).public_send(op, value)
    result.is_a?(Pptx::Length) ? result.emu : result
  rescue StandardError
    :error
  end

  def normalize_python(entry)
    entry["ok"] ? entry["value"] : :error
  end

  # python-pptx returns int where we may return Integer or Float for the same
  # quantity; compare numerics by value, everything else exactly.
  def comparable?(ours, theirs)
    return ours == theirs unless ours.is_a?(Numeric) && theirs.is_a?(Numeric)

    (ours - theirs).abs < 1e-9
  end

  describe "Ruby-side affordances" do
    it "accepts a Length wherever EMU integers are expected" do
      expect(st::ST_Coordinate.to_xml(Pptx.inches(1))).to eq("914400")
    end

    it "returns Length from length-typed attributes" do
      expect(st::ST_Coordinate.from_xml("914400")).to eq(Pptx.inches(1))
      expect(st::ST_TextSpacingPoint.from_xml("1800").pt).to eq(18.0)
    end

    it "exposes enumeration members as constants and validates against them" do
      aggregate_failures do
        expect(st::ST_BarDir::COL).to eq("col")
        expect(st::ST_BarDir.to_xml("bar")).to eq("bar")
        expect { st::ST_BarDir.to_xml("sideways") }.to raise_error(ArgumentError, /must be one of/)
      end
    end

    it "rounds half-to-even like Python, not half-up like Ruby" do
      # 0.000005 * 100000 == 0.5 exactly; Ruby's default round would give 1.
      expect(st::ST_Percentage.to_xml(0.000005)).to eq("0")
      expect(st::ST_Percentage.to_xml(0.000015)).to eq("2")
    end
  end
end
