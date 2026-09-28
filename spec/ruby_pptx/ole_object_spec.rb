# frozen_string_literal: true

RSpec.describe "OLE objects" do
  def self.images = File.expand_path("../fixtures/images", __dir__)
  def png = File.join(self.class.images, "png-96dpi.png")

  let(:deck) { Pptx::Presentation.new_default }
  let(:slide) { deck.slides.add(deck.slide_layouts[6]) }

  # Content is irrelevant to embedding; any bytes will do, and distinct ones
  # show each part holds its own file.
  around do |example|
    Dir.mktmpdir do |dir|
      @dir = dir
      example.run
    end
  end

  def object_file(name, content = "#{name} bytes")
    File.join(@dir, name).tap { |path| File.binwrite(path, content) }
  end

  def partnames = deck.part.package.parts.map { |part| part.partname.to_s }

  describe "embedding" do
    it "embeds an Office document as the document, related as a package" do
      frame = slide.shapes.add_ole_object(object_file("budget.xlsx"), prog_id: :xlsx, at: [0, 0])
      rel = slide.part.rels.find { |r| r.target_part.partname.to_s.include?("embeddings") }
      aggregate_failures do
        expect(partnames).to include("/ppt/embeddings/Microsoft_Excel_Sheet1.xlsx")
        expect(rel.reltype).to eq(Pptx::Opc::RELATIONSHIP_TYPE::PACKAGE)
        expect(frame.ole_format.prog_id).to eq("Excel.Sheet.12")
      end
    end

    it "names each Office type's part for what it is" do
      slide.shapes.add_ole_object(object_file("a.docx"), prog_id: :docx, at: [0, 0])
      slide.shapes.add_ole_object(object_file("b.pptx"), prog_id: :pptx, at: [0, 0])
      aggregate_failures do
        expect(partnames).to include("/ppt/embeddings/Microsoft_Word_Document1.docx")
        expect(partnames).to include("/ppt/embeddings/Microsoft_PowerPoint_Presentation1.pptx")
      end
    end

    it "embeds anything else as an opaque object, related as an OLE object" do
      frame = slide.shapes.add_ole_object(object_file("x.pdf"), prog_id: "AcroExch.Document", at: [0, 0])
      rel = slide.part.rels.find { |r| r.target_part.partname.to_s.include?("embeddings") }
      aggregate_failures do
        expect(partnames).to include("/ppt/embeddings/oleObject1.bin")
        expect(rel.reltype).to eq(Pptx::Opc::RELATIONSHIP_TYPE::OLE_OBJECT)
        expect(frame.ole_format.prog_id).to eq("AcroExch.Document")
      end
    end

    it "reads the embedded bytes back" do
      frame = slide.shapes.add_ole_object(object_file("budget.xlsx", "the workbook"), prog_id: :xlsx,
                                                                                      at: [0, 0])
      aggregate_failures do
        expect(frame.ole_format.blob).to eq("the workbook")
        expect(frame.ole_format).to be_show_as_icon
        expect(frame.shape_type.name).to eq(:EMBEDDED_OLE_OBJECT)
        expect(frame).to be_ole_object
      end
    end

    it "takes an open stream as well as a path" do
      frame = File.open(object_file("budget.xlsx", "streamed")) do |io|
        slide.shapes.add_ole_object(io, prog_id: :xlsx, at: [0, 0])
      end
      expect(frame.ole_format.blob).to eq("streamed")
    end

    it "refuses a ProgID symbol it does not know" do
      expect { slide.shapes.add_ole_object(object_file("a"), prog_id: :odt, at: [0, 0]) }
        .to raise_error(ArgumentError, /not a member of PROG_ID/)
    end

    # A linked object points at a file outside the presentation instead of
    # carrying one, which shows as the absence of p:embed.
    it "reports a linked object as linked" do
      frame = slide.shapes.add_ole_object(object_file("a.xlsx"), prog_id: :xlsx, at: [0, 0])
      frame.element.xpath(".//p:embed").each { |embed| embed.parent.remove(embed) }
      expect(frame.shape_type.name).to eq(:LINKED_OLE_OBJECT)
    end

    # A frame can hold a diagram or other content this library does not model.
    it "reports no shape type for a frame holding something it does not model" do
      frame = slide.shapes.add_ole_object(object_file("a.xlsx"), prog_id: :xlsx, at: [0, 0])
      frame.element.xpath("./a:graphic/a:graphicData").first.set("uri", "urn:something-else")
      aggregate_failures do
        expect(frame.shape_type).to be_nil
        expect(frame).not_to be_ole_object
      end
    end

    it "refuses ole_format on a frame holding something else" do
      table = slide.shapes.add_table(1, 1, at: [0, 0], size: [100, 100])
      expect { table.ole_format }.to raise_error(Pptx::Error, /does not contain an OLE object/)
    end
  end

  # PowerPoint stores an EMF icon as imageN.emf; python-pptx, asking Pillow,
  # stores it as .wmf -- see PORTING.md.
  describe "the default icons" do
    it "uses Office's icon for an Office document and a generic one otherwise" do
      xlsx = slide.shapes.add_ole_object(object_file("a.xlsx"), prog_id: :xlsx, at: [0, 0])
      other = slide.shapes.add_ole_object(object_file("b.bin"), prog_id: "Package", at: [0, 0])
      icon = lambda do |frame|
        r_id = frame.element.xpath(".//a:blip/@r:embed").first.value
        slide.part.related_part(r_id)
      end
      templates = File.expand_path("../../lib/ruby_pptx/templates", __dir__)
      aggregate_failures do
        expect(icon.call(xlsx).blob).to eq(File.binread(File.join(templates, "xlsx-icon.emf")))
        expect(icon.call(other).blob).to eq(File.binread(File.join(templates, "generic-icon.emf")))
        expect(icon.call(xlsx).partname.to_s).to end_with(".emf")
        expect(icon.call(xlsx).content_type).to eq("image/x-emf")
      end
    end
  end

  describe Pptx::Image do
    it "recognises an EMF and reads its size and resolution as Pillow does" do
      image = Pptx::Image.from_file(File.expand_path("../../lib/ruby_pptx/templates/pptx-icon.emf", __dir__))
      aggregate_failures do
        expect(image.format).to eq(:EMF)
        expect(image.size).to eq([101, 58])
        expect(image.dpi).to eq([93, 83])
      end
    end
  end

  describe "agreement with python-pptx" do
    before { require_oracle! }

    # A PNG icon keeps the comparison byte-for-byte; the default EMF icons
    # are named differently by design and checked above.
    it "embeds each kind of object identically" do
      files = { "xlsx" => object_file("a.xlsx"), "docx" => object_file("b.docx"),
                "pptx" => object_file("c.pptx"), "bin" => object_file("d.bin") }
      expect_same_package(
        python: <<~PY,
          from pptx.enum.shapes import PROG_ID
          from pptx.util import Inches
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          icon = #{png.inspect}
          slide.shapes.add_ole_object(#{files["xlsx"].inspect}, PROG_ID.XLSX, Inches(1), Inches(1),
                                      icon_file=icon)
          slide.shapes.add_ole_object(#{files["docx"].inspect}, PROG_ID.DOCX, Inches(3), Inches(1),
                                      Inches(2), Inches(1.5), icon_file=icon,
                                      icon_width=Inches(0.5), icon_height=Inches(0.4))
          slide.shapes.add_ole_object(#{files["pptx"].inspect}, PROG_ID.PPTX, 0, 0, icon_file=icon)
          slide.shapes.add_ole_object(#{files["bin"].inspect}, "Package", Inches(5), 0, icon_file=icon)
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          slide.shapes.add_ole_object(files["xlsx"], prog_id: :xlsx, at: [Pptx.inches(1), Pptx.inches(1)],
                                                     icon_file: png)
          slide.shapes.add_ole_object(files["docx"], prog_id: :docx, at: [Pptx.inches(3), Pptx.inches(1)],
                                                     size: [Pptx.inches(2), Pptx.inches(1.5)], icon_file: png,
                                                     icon_size: [Pptx.inches(0.5), Pptx.inches(0.4)])
          slide.shapes.add_ole_object(files["pptx"], prog_id: :pptx, at: [0, 0], icon_file: png)
          slide.shapes.add_ole_object(files["bin"], prog_id: "Package", at: [Pptx.inches(5), 0],
                                                    icon_file: png)
          prs.save(path)
        }
      )
    end

    # The whole-package comparison cannot cover the default icons, so the
    # divergence is pinned here instead: same bytes, different name.
    it "stores the default icon under the name PowerPoint uses, where python-pptx says wmf" do
      xlsx = object_file("a.xlsx")
      script = <<~PY
        import sys, pptx
        from pptx.enum.shapes import PROG_ID
        prs = pptx.Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_ole_object(#{xlsx.inspect}, PROG_ID.XLSX, 0, 0)
        part = [p for p in prs.part.package.iter_parts() if "media" in str(p.partname)][0]
        print(part.partname, part.content_type)
      PY
      out, err, status = Open3.capture3("python3", "-c", script)
      raise "python-pptx failed: #{err}" unless status.success?

      slide.shapes.add_ole_object(xlsx, prog_id: :xlsx, at: [0, 0])
      ours = deck.part.package.parts.find { |p| p.partname.to_s.include?("media") }
      aggregate_failures do
        expect(out.split).to eq(["/ppt/media/image1.wmf", "image/x-wmf"])
        expect([ours.partname.to_s, ours.content_type]).to eq(["/ppt/media/image1.emf", "image/x-emf"])
      end
    end

    it "reads an object python-pptx embedded" do
      result = Tempfile.create(["py-ole", ".pptx"]) do |file|
        file.close
        script = <<~PY
          import sys, pptx
          from pptx.enum.shapes import PROG_ID
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          slide.shapes.add_ole_object(#{object_file("a.docx", "doc bytes").inspect}, PROG_ID.DOCX, 0, 0)
          prs.save(sys.argv[1])
        PY
        _, err, status = Open3.capture3("python3", "-c", script, file.path)
        raise "python-pptx failed: #{err}" unless status.success?

        frame = Pptx::Presentation.open(file.path).slides[0].shapes[0]
        [frame.shape_type.name, frame.ole_format.prog_id, frame.ole_format.blob,
         frame.ole_format.show_as_icon?]
      end
      expect(result).to eq([:EMBEDDED_OLE_OBJECT, "Word.Document.12", "doc bytes", true])
    end
  end
end
