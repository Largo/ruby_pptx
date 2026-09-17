# frozen_string_literal: true

# Validates the differential harness itself: python-pptx writes a deck, and the
# Ruby and Python canonicalizers must produce identical manifests for it. If
# this fails, every differential result is untrustworthy.
RSpec.describe Pptx::Spec::Differential do
  before { require_oracle! }

  let(:script) do
    <<~PY
      prs = pptx.Presentation()
      slide = prs.slides.add_slide(prs.slide_layouts[1])
      slide.shapes.title.text = "Hello"
      slide.placeholders[1].text_frame.text = "from the oracle"
      prs.save(out)
    PY
  end

  it "agrees with the Python canonicalizer on a real package" do
    expected = python_pptx_manifest(script)

    actual = Dir.mktmpdir do |dir|
      path = File.join(dir, "oracle.pptx")
      out, err, status = Open3.capture3(
        "python3", "-c", "import sys,pptx; out=sys.argv[1]; exec(open(sys.argv[2]).read())",
        path, write_script(dir, script)
      )
      raise "oracle write failed: #{err}#{out}" unless status.success?

      ruby_pptx_manifest(path)
    end

    expect { compare_manifests(expected, actual, []) }.not_to raise_error
    expect(actual["entries"]).to include("ppt/presentation.xml", "ppt/slides/slide1.xml")
  end

  def write_script(dir, script)
    File.join(dir, "s.py").tap { |p| File.write(p, script) }
  end
end
