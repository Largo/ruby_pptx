# frozen_string_literal: true

RSpec.describe Pptx::Length do
  describe "construction from each unit" do
    {
      emu: [914_400, 914_400],
      inches: [1, 914_400],
      cm: [2.54, 914_400],
      mm: [25.4, 914_400],
      pt: [72, 914_400],
      centipoints: [7200, 914_400]
    }.each do |unit, (value, expected_emu)|
      it "converts #{value} #{unit} to #{expected_emu} EMU" do
        expect(described_class.public_send(unit, value).emu).to eq(expected_emu)
      end
    end

    it "truncates toward zero, matching python-pptx" do
      expect(described_class.inches(1.0000001).emu).to eq(914_400)
      expect(described_class.inches(-1.0000001).emu).to eq(-914_400)
    end
  end

  describe "conversion out" do
    subject(:length) { described_class.emu(914_400) }

    it { expect(length.inches).to eq(1.0) }
    it { expect(length.cm).to be_within(1e-9).of(2.54) }
    it { expect(length.mm).to be_within(1e-9).of(25.4) }
    it { expect(length.pt).to eq(72.0) }
    it { expect(length.centipoints).to eq(7200) }
    it { expect(length.to_i).to eq(914_400) }
  end

  it "compares against a bare Integer as EMU" do
    expect(described_class.inches(1)).to eq(914_400)
    expect(described_class.inches(1)).to be > described_class.cm(2)
    expect([described_class.pt(10), described_class.pt(1)].min).to eq(described_class.pt(1))
  end

  it "adds, subtracts, negates and scales, staying a Length" do
    aggregate_failures do
      expect(described_class.inches(1) + described_class.inches(2)).to eq(described_class.inches(3))
      expect(described_class.inches(3) - described_class.inches(1)).to eq(described_class.inches(2))
      expect(-described_class.inches(1)).to eq(described_class.inches(-1))
      expect(described_class.inches(1) * 2.5).to eq(described_class.inches(2.5))
      expect(described_class.inches(1) + described_class.inches(1)).to be_a(described_class)
    end
  end

  it "is frozen and usable as a Hash key" do
    expect(described_class.pt(1)).to be_frozen
    expect({ described_class.pt(1) => :a }[described_class.centipoints(100)]).to eq(:a)
  end

  it "coerces nil, Integer and Length; rejects anything else" do
    aggregate_failures do
      expect(described_class.coerce(nil)).to be_nil
      expect(described_class.coerce(127)).to eq(described_class.centipoints(1))
      expect { described_class.coerce("1in") }.to raise_error(TypeError, /Integer/)
    end
  end

  it "exposes module-level constructors" do
    expect(Pptx.inches(1)).to eq(Pptx.emu(914_400))
  end

  describe "opt-in core extensions" do
    it "adds unit methods to Numeric only when required" do
      require "ruby_pptx/core_ext"
      expect(1.inch).to eq(Pptx.inches(1))
      expect(2.54.cm).to eq(Pptx.inches(1))
      expect(72.points).to eq(Pptx.inches(1))
    end
  end
end
