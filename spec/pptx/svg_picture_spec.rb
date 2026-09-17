# frozen_string_literal: true

# SVG pictures are beyond python-pptx, which explicitly skips SVG as an
# unsupported image type, so there is no differential oracle. They are checked
# against the structure in the Microsoft [MS-ODRAWXML] specification, and by
# proving python-pptx can still open a deck containing one.
RSpec.describe "SVG pictures" do
  def self.images = File.expand_path("../fixtures/images", __dir__)
  def image(name) = File.join(self.class.images, name)

  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }

  def add_svg(**options)
    slide.shapes.add_picture(image("vector.svg"), at: [Pptx.inches(1), Pptx.inches(1)],
                                                  fallback: image("png-96dpi.png"), **options)
  end

  describe Pptx::Image do
    subject(:svg) { Pptx::Image.from_file(image("vector.svg")) }

    it "recognizes an SVG by its root element rather than a magic number" do
      aggregate_failures do
        expect(svg.format).to eq(:SVG)
        expect(svg).to be_vector
        expect(svg.ext).to eq("svg")
        expect(svg.content_type).to eq("image/svg+xml")
      end
    end

    it "recognizes an SVG with no XML declaration" do
      bare = Pptx::Image.from_blob(%(<svg xmlns="http://www.w3.org/2000/svg"/>))
      expect(bare.format).to eq(:SVG)
    end

    # An SVG is resolution-independent; asking it for pixels is a category
    # error, and silently guessing would size pictures wrongly.
    it "refuses to invent a native size" do
      expect { svg.native_size }
        .to raise_error(Pptx::Error, /no native pixel size/)
    end

    it "does not mistake a raster for a vector" do
      expect(Pptx::Image.from_file(image("png-96dpi.png"))).not_to be_vector
    end
  end

  describe "adding one" do
    it "embeds the fallback and hangs the SVG off the blip extension" do
      pic = add_svg
      blip = pic.element.blip
      aggregate_failures do
        expect(blip.embed).not_to be_nil
        expect(blip.svg_rId).not_to be_nil
        expect(blip.embed).not_to eq(blip.svg_rId)
      end
    end

    it "writes the extension exactly as the Microsoft specification gives it" do
      xml = add_svg.element.xml.gsub(/\s+/, " ")
      aggregate_failures do
        expect(xml).to include(%(<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">))
        expect(xml).to include(
          %(xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main")
        )
        expect(xml).to match(/<asvg:svgBlip[^>]*r:embed="rId\d+"/)
      end
    end

    it "stores both images as parts, each with its own content type" do
      add_svg
      partnames = presentation.part.package.parts.map { |p| p.partname.to_s }
      expect(partnames.grep(%r{/ppt/media/}).sort)
        .to eq(["/ppt/media/image1.png", "/ppt/media/image2.svg"])
    end

    # The SVG has no pixel size, so the fallback decides -- which is also what
    # keeps the two in agreement.
    it "sizes the picture from the fallback" do
      pic = add_svg
      aggregate_failures do
        expect(pic.width.inches).to be_within(1e-9).of(64 / 96.0)
        expect(pic.height.inches).to be_within(1e-9).of(48 / 96.0)
      end
    end

    it "still honours an explicit size" do
      pic = add_svg(width: Pptx.inches(4))
      aggregate_failures do
        expect(pic.width).to eq(Pptx.inches(4))
        expect(pic.height).to eq(Pptx.inches(3))
      end
    end

    it "refuses an SVG with no fallback, and says why" do
      expect { slide.shapes.add_picture(image("vector.svg"), at: [0, 0]) }
        .to raise_error(ArgumentError, /needs a raster fallback.*cannot rasterize/m)
    end

    it "refuses a fallback for a raster, which would be meaningless" do
      expect {
        slide.shapes.add_picture(image("png-96dpi.png"), at: [0, 0],
                                                         fallback: image("gif.gif"))
      }.to raise_error(ArgumentError, /only meaningful for a vector/)
    end

    it "leaves a plain raster picture without an extension list" do
      pic = slide.shapes.add_picture(image("png-96dpi.png"), at: [0, 0])
      aggregate_failures do
        expect(pic.element.blip.svg_rId).to be_nil
        expect(pic.element.xml).not_to include("extLst")
      end
    end
  end

  describe "round-tripping through python-pptx" do
    before { require_oracle! }

    # python-pptx has no SVG support at all -- it skips SVG as an unknown image
    # type -- but it must still open the file and leave the extension intact.
    it "survives being opened and saved by python-pptx" do
      add_svg
      Tempfile.create(["svg", ".pptx"]) do |file|
        file.close
        presentation.save(file.path)

        script = <<~PY
          import sys, zipfile, re, pptx
          prs = pptx.Presentation(#{file.path.inspect})
          prs.save(#{file.path.inspect})
          xml = zipfile.ZipFile(#{file.path.inspect}).read("ppt/slides/slide1.xml").decode()
          m = re.search(r"(?s)<a:blip.*?</a:blip>", xml)
          sys.stdout.write(m.group(0) if m else "")
        PY
        out, err, status = Open3.capture3("python3", "-c", script)
        raise "round-trip failed: #{err}" unless status.success?

        aggregate_failures do
          expect(out).to include("svgBlip")
          expect(out).to include("{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
        end

        reopened = Pptx::Presentation.open(file.path)
        expect(reopened.slides.first.shapes.first.element.blip.svg_rId).not_to be_nil
      end
    end
  end
end
