# frozen_string_literal: true

RSpec.describe "picture cropping, masking and outline" do
  def self.images = File.expand_path("../fixtures/images", __dir__)
  def images = self.class.images
  def png = File.join(images, "png-96dpi.png")

  let(:deck) { Pptx::Presentation.new_default }
  let(:slide) { deck.slides.add(deck.slide_layouts[6]) }
  let(:picture) { slide.shapes.add_picture(png, at: [0, 0]) }

  describe Pptx::Picture do
    it "reads back the image it shows" do
      image = picture.image
      aggregate_failures do
        expect(image).to be_a(Pptx::Image)
        expect(image.blob).to eq(File.binread(png))
        expect(image.content_type).to eq("image/png")
        expect(image.size).to eq([64, 48])
      end
    end

    # Two different images, so the read-back cannot pass by returning a
    # fixed one: each picture must resolve its own relationship.
    it "reads back each picture's own image" do
      wide = File.join(images, "wide-97x31.png")
      second = slide.shapes.add_picture(wide, at: [0, 0])
      aggregate_failures do
        expect(picture.image.blob).to eq(File.binread(png))
        expect(second.image.blob).to eq(File.binread(wide))
        expect(second.image.size).to eq([97, 31])
      end
    end

    it "reports no cropping until some is set" do
      expect([picture.crop_left, picture.crop_top, picture.crop_right, picture.crop_bottom])
        .to eq([0.0, 0.0, 0.0, 0.0])
    end

    it "crops each side independently" do
      picture.crop_left = 0.1
      picture.crop_top = 0.2
      picture.crop_right = 0.3
      picture.crop_bottom = 0.4
      expect([picture.crop_left, picture.crop_top, picture.crop_right, picture.crop_bottom])
        .to eq([0.1, 0.2, 0.3, 0.4])
    end

    # PowerPoint uses negative crops to pad, and crops above 1.0 are legal.
    it "accepts negative and oversized crops" do
      picture.crop_left = -0.25
      picture.crop_right = 1.5
      expect([picture.crop_left, picture.crop_right]).to eq([-0.25, 1.5])
    end

    it "starts masked by a rectangle and can be masked by another shape" do
      aggregate_failures do
        expect(picture.auto_shape_type.name).to eq(:RECTANGLE)
        picture.auto_shape_type = :oval
        expect(picture.auto_shape_type.name).to eq(:OVAL)
      end
    end

    it "outlines the picture" do
      picture.line.width = Pptx.pt(3)
      expect(picture.line.width).to eq(Pptx.pt(3))
    end

    it "refuses an image read-back when the image is linked, not embedded" do
      picture.element.blip.node.remove_attribute("embed")
      expect { picture.image }.to raise_error(Pptx::Error, /no embedded image/)
    end
  end

  describe Pptx::Movie do
    let(:movie) do
      slide.shapes.add_movie(File.expand_path("../fixtures/movie.mp4", __dir__),
                             at: [0, 0], size: [Pptx.inches(4), Pptx.inches(3)],
                             content_type: "video/mp4")
    end

    it "is a movie and not a picture" do
      aggregate_failures do
        expect(movie).to be_a(Pptx::BasePicture)
        expect(movie).not_to be_a(Pptx::Picture)
        expect(movie.media_type.name).to eq(:MOVIE)
        expect(movie).not_to respond_to(:image)
      end
    end

    it "reads back its poster frame" do
      expect(movie.poster_frame).to be_a(Pptx::Image)
    end

    it "crops like a picture" do
      movie.crop_top = 0.1
      expect(movie.crop_top).to eq(0.1)
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "crops a picture identically" do
      expect_same_package(
        python: <<~PY,
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          pic = slide.shapes.add_picture(#{png.inspect}, 0, 0)
          pic.crop_left = 0.1
          pic.crop_top = 0.2
          pic.crop_right = -0.125
          pic.crop_bottom = 1.5
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          pic = slide.shapes.add_picture(png, at: [0, 0])
          pic.crop_left = 0.1
          pic.crop_top = 0.2
          pic.crop_right = -0.125
          pic.crop_bottom = 1.5
          prs.save(path)
        }
      )
    end

    it "masks and outlines a picture identically" do
      expect_same_package(
        python: <<~PY,
          from pptx.enum.shapes import MSO_SHAPE
          from pptx.util import Pt
          from pptx.dml.color import RGBColor
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          pic = slide.shapes.add_picture(#{png.inspect}, 0, 0)
          pic.auto_shape_type = MSO_SHAPE.OVAL
          pic.line.width = Pt(3)
          pic.line.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          pic = slide.shapes.add_picture(png, at: [0, 0])
          pic.auto_shape_type = :oval
          pic.line.width = Pptx.pt(3)
          pic.line.color.rgb = Pptx::RGBColor["1F497D"]
          prs.save(path)
        }
      )
    end

    # Reading a crop python-pptx wrote checks the conversion the other way.
    it "reads crops and an image written by python-pptx" do
      result = Tempfile.create(["py-crop", ".pptx"]) do |file|
        file.close
        script = <<~PY
          import sys, pptx
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          pic = slide.shapes.add_picture(#{png.inspect}, 0, 0)
          pic.crop_left = 0.33333
          pic.crop_bottom = 0.5
          prs.save(sys.argv[1])
        PY
        _, err, status = Open3.capture3("python3", "-c", script, file.path)
        raise "python-pptx failed: #{err}" unless status.success?

        pic = Pptx::Presentation.open(file.path).slides[0].shapes[0]
        [pic.crop_left, pic.crop_bottom, pic.image.sha1]
      end
      expect(result).to eq([0.33333, 0.5, Digest::SHA1.hexdigest(File.binread(png))])
    end
  end
end
