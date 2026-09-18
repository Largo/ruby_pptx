# frozen_string_literal: true

RSpec.describe "what the gem ships" do
  def gem_root = File.expand_path("../..", __dir__)

  # A gem owns the require path that matches its name and nothing else. This
  # one used to ship lib/pptx.rb and lib/pptx/**, which the `pptx` gem on
  # RubyGems already owns -- six files collided, `version.rb` among them. With
  # both installed, `require "ruby_pptx"` loaded whichever lib/pptx.rb came
  # first on $LOAD_PATH, so Pptx::VERSION could simply not exist.
  it "keeps every file under the path matching the gem name" do
    files = Dir.glob("lib/**/*", base: gem_root).reject { |f| File.directory?(File.join(gem_root, f)) }
    strays = files.reject { |f| f == "lib/ruby_pptx.rb" || f.start_with?("lib/ruby_pptx/") }
    expect(strays).to eq([])
  end

  it "is requirable by the name it is published under" do
    expect(File.exist?(File.join(gem_root, "lib/ruby_pptx.rb"))).to be(true)
  end

  # The templates are data, not code, so the .rb glob does not cover them.
  it "packages the templates it reads at runtime" do
    spec = Gem::Specification.load(File.join(gem_root, "ruby_pptx.gemspec"))
    aggregate_failures do
      expect(spec.name).to eq("ruby_pptx")
      expect(spec.files).to include("lib/ruby_pptx/templates/default.pptx")
      expect(spec.files).to include("lib/ruby_pptx/templates/slideMaster.xml")
      expect(spec.files).to include("lib/ruby_pptx/templates/media-speaker.png")
    end
  end

  # Every .rb the gemspec claims must actually load from a clean interpreter,
  # which is what catches a require left pointing at the old path.
  it "loads from a clean interpreter with only lib on the path" do
    out, err, status = Open3.capture3("ruby", "-I", File.join(gem_root, "lib"),
                                      "-e", 'require "ruby_pptx"; print Pptx::VERSION')
    raise "load failed: #{err}" unless status.success?

    expect(out).to eq(Pptx::VERSION)
  end
end
