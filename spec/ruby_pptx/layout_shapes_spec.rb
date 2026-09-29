# frozen_string_literal: true

require "stringio"

# Shapes that are not placeholders, drawn on a layout or a master: every slide
# built on it shows them, which is how a corporate template carries its logo,
# bands and rules.
RSpec.describe "drawing on a layout or master" do
  let(:presentation) { Pptx::Presentation.new_default }
  let(:master) { presentation.slide_masters.first }
  let(:layout) { presentation.slide_layouts["Blank"] }

  def images = File.expand_path("../fixtures/images", __dir__)
  def inches(*values) = values.map { |v| Pptx.inches(v) }

  def reopen
    io = StringIO.new
    presentation.save(io)
    Pptx::Presentation.open(StringIO.new(io.string))
  end

  { "layout" => :layout, "master" => :master }.each do |label, owner|
    describe "on a #{label}" do
      subject(:shapes) { public_send(owner).shapes }

      it "adds an auto shape, a text box and a connector" do
        band = shapes.add_shape(:rectangle, at: inches(0, 0), size: inches(10, 1))
        box = shapes.add_textbox(at: inches(1, 6), size: inches(3, 0.5))
        box.text_frame.text = "www.alos.ch"
        rule = shapes.add_connector(:straight, begin_at: inches(0.9, 1.75), end_at: inches(9, 1.75))
        aggregate_failures do
          expect(band).to be_a(Pptx::Shape)
          expect(box.text_frame.text).to eq("www.alos.ch")
          expect([rule.begin_x, rule.end_x, rule.begin_y]).to eq(inches(0.9, 9, 1.75))
          expect(shapes.to_a.last(3).map(&:shape_id)).to eq([band, box, rule].map(&:shape_id))
        end
      end

      it "adds an SVG picture with its raster fallback, related from its own part" do
        svg = File.join(images, "vector.svg")
        pic = shapes.add_picture(svg, at: inches(1, 1), fallback: File.join(images, "png-96dpi.png"))
        r_ids = pic.element.to_xml.scan(/r:embed="([^"]+)"/).flatten
        parts = r_ids.map { |r_id| public_send(owner).part.related_part(r_id) }
        aggregate_failures do
          expect(pic).to be_a(Pptx::Picture)
          expect(pic.element.to_xml).to include("svgBlip")
          expect(parts.map(&:content_type)).to contain_exactly("image/png", "image/svg+xml")
        end
      end

      it "adds a freeform and a group" do
        line = shapes.add_freeform(start_x: 0, start_y: 0, scale: Pptx.inches(1).emu / 100.0,
                                   close: false) { |f| f.line_to(300, 0) }
        group = shapes.add_group_shape([line])
        expect(group.shapes.to_a.map(&:shape_id)).to eq([line.shape_id])
      end

      it "keeps them out of the placeholder view" do
        placeholders = public_send(owner).placeholders
        before = placeholders.size
        shapes.add_shape(:rectangle, at: inches(0, 0), size: inches(1, 1))
        aggregate_failures do
          expect(placeholders.size).to eq(before)
          expect { placeholders.add_shape(:rectangle, at: inches(0, 0), size: inches(1, 1)) }
            .to raise_error(NoMethodError)
        end
      end

      it "survives a save" do
        shapes.add_shape(:rectangle, at: inches(0, 8), size: inches(10, 2)).name = "Band"
        shapes.add_picture(File.join(images, "png-96dpi.png"), at: inches(1, 1)).name = "Logo"
        reopened = reopen
        target = owner == :layout ? reopened.slide_layouts["Blank"] : reopened.slide_masters.first
        names = target.shapes.map(&:name)
        aggregate_failures do
          expect(names).to include("Band", "Logo")
          expect(target.shapes.find { |s| s.name == "Logo" }.image.blob).not_to be_empty
        end
      end
    end
  end

  it "does not let a slide built on the layout see the layout's shapes as its own" do
    layout.shapes.add_shape(:rectangle, at: inches(0, 0), size: inches(1, 1))
    slide = presentation.slides.add(layout)
    expect(slide.shapes.size).to eq(0)
  end
end
