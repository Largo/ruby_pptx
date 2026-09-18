# frozen_string_literal: true

RSpec.describe "movies" do
  def self.fixtures = File.expand_path("../fixtures", __dir__)
  def movie_file = File.join(self.class.fixtures, "movie.mp4")
  def poster_file = File.join(self.class.fixtures, "images/png-96dpi.png")

  let(:presentation) { Pptx::Presentation.new_default }
  let(:slide) { presentation.slides.add(presentation.slide_layouts["Blank"]) }

  def add_movie(**options)
    slide.shapes.add_movie(movie_file, at: [Pptx.inches(1), Pptx.inches(1)],
                                       size: [Pptx.inches(4), Pptx.inches(3)],
                                       content_type: "video/mp4", **options)
  end

  describe Pptx::Video do
    it "takes its type from the caller rather than the bytes" do
      video = described_class.from_file(movie_file, "video/mp4")
      aggregate_failures do
        expect(video.content_type).to eq("video/mp4")
        expect(video.ext).to eq("mp4")
        expect(video.filename).to eq("movie.mp4")
      end
    end

    # PowerPoint plays "video/unknown" anyway, which is why it is the default.
    it "falls back to an unknown type, with a generic extension" do
      video = described_class.from_blob("bytes", nil)
      aggregate_failures do
        expect(video.content_type).to eq(Pptx::Opc::CONTENT_TYPE::VIDEO)
        expect(video.ext).to eq("vid")
        expect(video.filename).to eq("movie.vid")
      end
    end

    it "derives the extension from the type when there is no filename" do
      aggregate_failures do
        expect(described_class.from_blob("x", "video/mp4").ext).to eq("mp4")
        expect(described_class.from_blob("x", "video/x-ms-wmv").ext).to eq("wmv")
        expect(described_class.from_blob("x", "video/avi").ext).to eq("avi")
      end
    end

    it "prefers the original filename's extension" do
      expect(described_class.from_file(movie_file, "video/unknown").ext).to eq("mp4")
    end

    it "reads from a stream as well as a path" do
      File.open(movie_file, "rb") do |io|
        video = described_class.from_file(io, "video/mp4")
        aggregate_failures do
          expect(video.filename).to eq("movie.mp4")
          expect(video.blob.bytesize).to eq(File.size(movie_file))
        end
      end
    end
  end

  describe "adding one" do
    subject(:movie) { add_movie }

    it "is a movie shape named after the video" do
      aggregate_failures do
        expect(movie).to be_a(Pptx::Movie)
        expect(movie.shape_type).to eq(Pptx::Enum::MSO_SHAPE_TYPE::MEDIA)
        expect(movie.name).to eq("movie.mp4")
      end
    end

    it "uses the size it was given, since a video has none to read" do
      aggregate_failures do
        expect(movie.width).to eq(Pptx.inches(4))
        expect(movie.height).to eq(Pptx.inches(3))
      end
    end

    it "stores the video as a media part and the poster as an image part" do
      movie
      partnames = presentation.part.package.parts.map { |p| p.partname.to_s }
      aggregate_failures do
        expect(partnames).to include("/ppt/media/media1.mp4")
        expect(partnames.grep(%r{/ppt/media/image\d+\.png}).size).to eq(1)
      end
    end

    # PowerPoint has embedded media two ways over the years, and writes both
    # so either era can find the same part.
    it "relates the media part twice, as media and as video" do
      movie
      reltypes = slide.part.rels.map(&:reltype)
      aggregate_failures do
        expect(reltypes).to include(Pptx::Opc::RELATIONSHIP_TYPE::MEDIA)
        expect(reltypes).to include(Pptx::Opc::RELATIONSHIP_TYPE::VIDEO)
        media = slide.part.rels.select do |r|
          [Pptx::Opc::RELATIONSHIP_TYPE::MEDIA,
           Pptx::Opc::RELATIONSHIP_TYPE::VIDEO].include?(r.reltype)
        end
        expect(media.map(&:target_part).uniq.size).to eq(1)
      end
    end

    it "writes the legacy link, the modern embed and the click action" do
      xml = movie.element.xml.gsub(/\s+/, " ")
      aggregate_failures do
        expect(xml).to match(%r{<a:videoFile r:link="rId\d+"/>})
        expect(xml).to include('<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">')
        expect(xml).to match(/<p14:media[^>]*r:embed="rId\d+"/)
        expect(xml).to include('action="ppaction://media"')
      end
    end

    # Without an entry in the slide's timing tree PowerPoint shows the poster
    # frame but no play controls.
    it "registers the movie in the slide's timing tree" do
      movie
      xml = slide.element.xml.gsub(/\s+/, " ")
      aggregate_failures do
        expect(xml).to include("<p:timing>")
        expect(xml).to include(%(<p:spTgt spid="#{movie.shape_id}"/>))
      end
    end

    it "uses the supplied poster frame when there is one" do
      with_poster = add_movie(poster_frame: poster_file)
      poster = presentation.part.package.parts
                           .find { |p| p.partname.to_s.start_with?("/ppt/media/image") }
      aggregate_failures do
        expect(with_poster).to be_a(Pptx::Movie)
        expect(poster.blob).to eq(File.binread(poster_file))
      end
    end

    it "stores one media part however many times the same video is added" do
      add_movie
      slide.shapes.add_movie(movie_file, at: [0, 0], size: [Pptx.inches(1), Pptx.inches(1)],
                                         content_type: "video/mp4")
      media = presentation.part.package.parts.map { |p| p.partname.to_s }
                          .grep(%r{/ppt/media/media})
      expect(media.size).to eq(1)
    end

    it "numbers timing nodes so a second movie does not collide" do
      first = add_movie
      second = slide.shapes.add_movie(movie_file, at: [0, 0],
                                                  size: [Pptx.inches(1), Pptx.inches(1)],
                                                  content_type: "video/mp4")
      ids = slide.element.xpath("//p:cTn/@id").map { |a| a.value.to_i }
      aggregate_failures do
        expect(ids.uniq).to eq(ids)
        expect(slide.element.xml).to include(%(<p:spTgt spid="#{first.shape_id}"/>))
        expect(slide.element.xml).to include(%(<p:spTgt spid="#{second.shape_id}"/>))
      end
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    it "adds a movie identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          slide.shapes.add_movie(#{movie_file.inspect}, Inches(1), Inches(1),
                                 Inches(4), Inches(3), mime_type="video/mp4")
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          slide.shapes.add_movie(movie_file, at: [Pptx.inches(1), Pptx.inches(1)],
                                             size: [Pptx.inches(4), Pptx.inches(3)],
                                             content_type: "video/mp4")
          prs.save(path)
        }
      )
    end

    it "adds a movie with a poster frame identically to python-pptx" do
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          slide.shapes.add_movie(#{movie_file.inspect}, Inches(1), Inches(1),
                                 Inches(4), Inches(3),
                                 poster_frame_image=#{poster_file.inspect},
                                 mime_type="video/mp4")
          slide.shapes.add_movie(#{movie_file.inspect}, Inches(1), Inches(5),
                                 Inches(2), Inches(1), mime_type="video/mp4")
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          slide.shapes.add_movie(movie_file, at: [Pptx.inches(1), Pptx.inches(1)],
                                             size: [Pptx.inches(4), Pptx.inches(3)],
                                             poster_frame: poster_file,
                                             content_type: "video/mp4")
          slide.shapes.add_movie(movie_file, at: [Pptx.inches(1), Pptx.inches(5)],
                                             size: [Pptx.inches(2), Pptx.inches(1)],
                                             content_type: "video/mp4")
          prs.save(path)
        }
      )
    end
  end
end
