# frozen_string_literal: true

RSpec.describe "filling placeholders" do
  def self.images = File.expand_path("../fixtures/images", __dir__)
  def images = self.class.images

  # 4:3, the same ratio as the template's picture placeholder, so nothing is
  # cropped; wide and tall ones are needed to exercise the cropping at all.
  let(:same_ratio) { File.join(images, "png-96dpi.png") }
  let(:wide) { File.join(images, "wide-97x31.png") }
  let(:tall) { File.join(images, "tall-23x89.png") }

  let(:deck) { Pptx::Presentation.new_default }

  def picture_slide = deck.slides.add(deck.slide_layouts["Picture with Caption"])

  # The default template has no table or chart placeholder, so the specs that
  # need one build a layout with them and both libraries open the result.
  def with_fillable_template
    Tempfile.create(["fillable", ".pptx"]) do |file|
      file.close
      prs = Pptx::Presentation.new_default
      prs.slide_masters[0].slide_layouts.add("Fillable", type: "obj") do |layout|
        layout.placeholders.add(:title, at: [Pptx.inches(0.5), Pptx.inches(0.3)],
                                        size: [Pptx.inches(9), Pptx.inches(1)])
        layout.placeholders.add(:table, idx: 1, at: [Pptx.inches(0.5), Pptx.inches(1.5)],
                                        size: [Pptx.inches(4.25), Pptx.inches(3)])
        layout.placeholders.add(:chart, idx: 2, at: [Pptx.inches(5), Pptx.inches(1.5)],
                                        size: [Pptx.inches(4.5), Pptx.inches(5)])
      end
      prs.save(file.path)
      yield file.path
    end
  end

  describe "which proxy a slide placeholder gets" do
    it "types placeholders by what they can be filled with" do
      with_fillable_template do |path|
        prs = Pptx::Presentation.open(path)
        slide = prs.slides.add(prs.slide_layouts["Fillable"])
        aggregate_failures do
          expect(slide.placeholders[1]).to be_a(Pptx::TablePlaceholder)
          expect(slide.placeholders[2]).to be_a(Pptx::ChartPlaceholder)
          expect(picture_slide.placeholders[1]).to be_a(Pptx::PicturePlaceholder)
        end
      end
    end

    # An object placeholder takes text in PowerPoint and has no insert_*.
    it "leaves other placeholders as plain slide placeholders" do
      slide = deck.slides.add(deck.slide_layouts["Title and Content"])
      aggregate_failures do
        expect(slide.placeholders[1].class).to eq(Pptx::SlidePlaceholder)
        expect(slide.placeholders[1]).not_to respond_to(:insert_picture)
      end
    end
  end

  describe Pptx::PicturePlaceholder do
    it "is replaced by a picture that is still a placeholder" do
      slide = picture_slide
      original = slide.placeholders[1]
      picture = original.insert_picture(same_ratio)
      aggregate_failures do
        expect(picture).to be_a(Pptx::PlaceholderPicture)
        expect(picture).to be_a(Pptx::Picture)
        expect(picture).to be_placeholder
        expect(picture.shape_type.name).to eq(:PLACEHOLDER)
        expect(picture.shape_id).to eq(original.shape_id)
        expect(picture.name).to eq(original.name)
        expect(slide.placeholders[1]).to be_a(Pptx::PlaceholderPicture)
        expect(slide.shapes.size).to eq(3)
      end
    end

    # No `a:xfrm`: the picture keeps taking its place from the layout.
    it "inherits its geometry rather than carrying any" do
      picture = picture_slide.placeholders[1].insert_picture(same_ratio)
      layout_ph = deck.slide_layouts["Picture with Caption"].placeholders.by_idx(1)
      aggregate_failures do
        expect(picture.element.xpath("./p:spPr/a:xfrm")).to be_empty
        expect([picture.left, picture.top, picture.width, picture.height])
          .to eq([layout_ph.left, layout_ph.top, layout_ph.width, layout_ph.height])
      end
    end

    it "crops nothing when the image already has the placeholder's shape" do
      picture = picture_slide.placeholders[1].insert_picture(same_ratio)
      expect([picture.crop_left, picture.crop_top, picture.crop_right, picture.crop_bottom])
        .to eq([0.0, 0.0, 0.0, 0.0])
    end

    it "crops a wide image equally from left and right" do
      picture = picture_slide.placeholders[1].insert_picture(wide)
      aggregate_failures do
        expect(picture.crop_left).to be > 0
        expect(picture.crop_left).to eq(picture.crop_right)
        expect([picture.crop_top, picture.crop_bottom]).to eq([0.0, 0.0])
      end
    end

    it "crops a tall image equally from top and bottom" do
      picture = picture_slide.placeholders[1].insert_picture(tall)
      aggregate_failures do
        expect(picture.crop_top).to be > 0
        expect(picture.crop_top).to eq(picture.crop_bottom)
        expect([picture.crop_left, picture.crop_right]).to eq([0.0, 0.0])
      end
    end

    # The visible part must have the placeholder's aspect ratio exactly.
    it "leaves exactly the placeholder's aspect ratio visible" do
      picture = picture_slide.placeholders[1].insert_picture(wide)
      visible = 97 * (1 - picture.crop_left - picture.crop_right)
      expect(visible / 31.0).to be_within(0.001).of(picture.width.emu.to_f / picture.height.emu)
    end
  end

  describe Pptx::TablePlaceholder do
    it "is replaced by a table at the placeholder's position and width" do
      with_fillable_template do |path|
        prs = Pptx::Presentation.open(path)
        slide = prs.slides.add(prs.slide_layouts["Fillable"])
        frame = slide.placeholders[1].insert_table(3, 4)
        aggregate_failures do
          expect(frame).to be_a(Pptx::PlaceholderGraphicFrame)
          expect(frame).to be_placeholder
          expect(frame.table.rows.size).to eq(3)
          expect(frame.table.columns.size).to eq(4)
          expect(frame.left).to eq(Pptx.inches(0.5))
          expect(frame.width).to eq(Pptx.inches(4.25))
          # Height follows the rows, not the placeholder.
          expect(frame.height).to eq(Pptx.emu(3 * 370_840))
        end
      end
    end
  end

  describe Pptx::ChartPlaceholder do
    it "is replaced by a chart filling the placeholder" do
      with_fillable_template do |path|
        prs = Pptx::Presentation.open(path)
        slide = prs.slides.add(prs.slide_layouts["Fillable"])
        data = Pptx::ChartData.new
        data.categories = %w[East West]
        data.add_series("Q1", [1, 2])
        frame = slide.placeholders[2].insert_chart(:column_clustered, data)
        aggregate_failures do
          expect(frame).to be_a(Pptx::PlaceholderGraphicFrame)
          expect(frame.chart.plot_type).to eq(:BAR)
          expect([frame.left, frame.top, frame.width, frame.height])
            .to eq([Pptx.inches(5), Pptx.inches(1.5), Pptx.inches(4.5), Pptx.inches(5)])
        end
      end
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    %w[png-96dpi.png wide-97x31.png tall-23x89.png].each do |name|
      it "fills a picture placeholder with #{name} identically" do
        image = File.join(images, name)
        expect_same_package(
          python: <<~PY,
            prs = pptx.Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[8])
            slide.placeholders[1].insert_picture(#{image.inspect})
            prs.save(out)
          PY
          ruby: lambda { |path|
            prs = Pptx::Presentation.new_default
            slide = prs.slides.add(prs.slide_layouts["Picture with Caption"])
            slide.placeholders[1].insert_picture(image)
            prs.save(path)
          }
        )
      end
    end

    it "fills table and chart placeholders identically" do
      with_fillable_template do |template|
        expect_same_package(
          python: <<~PY,
            from pptx.chart.data import CategoryChartData
            from pptx.enum.chart import XL_CHART_TYPE
            prs = pptx.Presentation(#{template.inspect})
            slide = prs.slides.add_slide(prs.slide_layouts.get_by_name("Fillable"))
            slide.shapes.title.text = "Filled"
            table = slide.placeholders[1].insert_table(3, 4).table
            table.cell(0, 0).text = "Region"
            cd = CategoryChartData()
            cd.categories = ["East", "West"]
            cd.add_series("Q1", (1.5, 2.25))
            slide.placeholders[2].insert_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, cd)
            prs.save(out)
          PY
          ruby: lambda { |path|
            prs = Pptx::Presentation.open(template)
            slide = prs.slides.add(prs.slide_layouts["Fillable"])
            slide.shapes.title.text = "Filled"
            table = slide.placeholders[1].insert_table(3, 4).table
            table.cell(0, 0).text = "Region"
            data = Pptx::ChartData.new
            data.categories = %w[East West]
            data.add_series("Q1", [1.5, 2.25])
            slide.placeholders[2].insert_chart(:column_clustered, data)
            prs.save(path)
          },
          # Written by XlsxWriter on one side and directly on the other; the
          # chart spec checks its contents separately.
          ignore: ["ppt/embeddings/Microsoft_Excel_Sheet1.xlsx"]
        )
      end
    end

    it "reads a placeholder python-pptx filled" do
      result = Tempfile.create(["py-ph", ".pptx"]) do |file|
        file.close
        script = <<~PY
          import sys, pptx
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[8])
          slide.placeholders[1].insert_picture(#{wide.inspect})
          prs.save(sys.argv[1])
        PY
        _, err, status = Open3.capture3("python3", "-c", script, file.path)
        raise "python-pptx failed: #{err}" unless status.success?

        picture = Pptx::Presentation.open(file.path).slides[0].placeholders[1]
        [picture.class, picture.crop_left.round(5), picture.width]
      end
      expect(result).to eq([Pptx::PlaceholderPicture, 0.28694, Pptx.inches(6)])
    end
  end
end
