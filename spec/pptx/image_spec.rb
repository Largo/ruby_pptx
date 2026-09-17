# frozen_string_literal: true

require "json"
require "open3"

RSpec.describe Pptx::Image do
  def self.fixture_dir = File.expand_path("../fixtures/images", __dir__)
  def fixture_dir = self.class.fixture_dir
  def fixtures = Dir[File.join(fixture_dir, "*")].sort

  def image_oracle_path = File.expand_path("../../tools/image_oracle.py", __dir__)

  def pillow_view(paths)
    out, err, status = Open3.capture3("python3", image_oracle_path, stdin_data: JSON.dump(paths))
    raise "image oracle failed (#{status.exitstatus}): #{err}" if out.empty?

    JSON.parse(out)
  end

  describe "agreement with Pillow" do
    before { skip "python-pptx not importable" unless Pptx::Spec::Differential.oracle_available? }

    # This parser exists to avoid depending on an image library, so it is
    # checked against the one it replaces, for every format we support.
    it "reads format, pixel size and dpi exactly as python-pptx does" do
      expected = pillow_view(fixtures)

      mismatches = fixtures.filter_map do |path|
        theirs = expected[path]
        ours = described_class.from_file(path)
        actual = {
          "format" => ours.format.to_s, "size" => ours.size,
          "dpi" => ours.dpi, "ext" => ours.ext,
          "content_type" => ours.content_type, "sha1" => ours.sha1
        }
        next if actual == theirs

        "#{File.basename(path)}: ours=#{actual.inspect} theirs=#{theirs.inspect}"
      end

      expect(mismatches).to be_empty, -> { mismatches.join("\n") }
    end
  end

  describe "resolution handling" do
    it "falls back to 72 dpi when the file does not say" do
      expect(described_class.from_file(File.join(fixture_dir, "gif.gif")).dpi).to eq([72, 72])
    end

    # A BMP header almost always carries 3780 pixels-per-metre rather than
    # zero, so a BMP reports about 96 dpi rather than the 72 default.
    it "reads a BMP's pixels-per-metre rather than defaulting" do
      expect(described_class.from_file(File.join(fixture_dir, "bmp.bmp")).dpi).to eq([96, 96])
    end

    it "rejects an implausible dpi in favour of the default" do
      image = described_class.from_file(File.join(fixture_dir, "png-96dpi.png"))
      allow(image).to receive(:header).and_return(
        { format: :PNG, width: 10, height: 10, horz_dpi: 99_999, vert_dpi: 0 }
      )
      expect(image.dpi).to eq([72, 72])
    end
  end

  describe "native size" do
    it "converts pixels to EMU through the resolution" do
      image = described_class.from_file(File.join(fixture_dir, "png-96dpi.png"))
      width, height = image.native_size
      aggregate_failures do
        expect(width.inches).to be_within(1e-9).of(64 / 96.0)
        expect(height.inches).to be_within(1e-9).of(48 / 96.0)
      end
    end
  end

  describe "identity" do
    it "derives the extension from the bytes, not the filename" do
      blob = File.binread(File.join(fixture_dir, "png-96dpi.png"))
      expect(described_class.from_blob(blob, "actually-a-lie.jpg").ext).to eq("png")
    end

    it "remembers a filename from a path but not from a stream" do
      path = File.join(fixture_dir, "gif.gif")
      aggregate_failures do
        expect(described_class.from_file(path).filename).to eq("gif.gif")
        File.open(path, "rb") { |io| expect(described_class.from_file(io).filename).to be_nil }
      end
    end

    it "raises on bytes that are not a recognized image" do
      expect { described_class.from_blob("not an image").format }
        .to raise_error(Pptx::Error, /unrecognized image format/)
    end
  end
end
