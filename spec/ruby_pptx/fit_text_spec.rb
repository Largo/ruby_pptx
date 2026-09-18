# frozen_string_literal: true

require "json"
require "open3"

RSpec.describe "fitting text to a shape" do
  # Any TrueType font will do -- both sides of the comparison are handed the
  # same file. DejaVu is the one CI installs.
  SANS = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
  SERIF = "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"
  # An OpenType/CFF font, which has no glyf table.
  CFF = "/usr/share/fonts/opentype/urw-base35/URWBookman-Light.otf"
  # A monospace font, whose hmtx table is much shorter than its glyph count.
  MONO = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"

  # Same bargain as the python-pptx oracle: missing locally is a skip, but in
  # CI -- where the workflow installs these -- it is a failure, so the specs
  # cannot quietly stop running.
  def font_or_skip(path)
    return path if File.exist?(path)
    raise "#{path} is not installed and REQUIRE_ORACLE is set" if ENV["REQUIRE_ORACLE"]

    skip "#{path} is not installed"
  end

  describe Pptx::FontMetrics do
    subject(:metrics) { described_class.open(font_or_skip(SANS)) }

    it "reads the em square from the font" do
      expect(metrics.units_per_em).to eq(2048)
    end

    it "measures wider text as wider" do
      narrow = metrics.text_extents("iiiii", 18).first
      wide = metrics.text_extents("WWWWW", 18).first
      expect(wide).to be > narrow
    end

    it "scales linearly with point size" do
      at_12 = metrics.text_extents("Hamburgefonstiv", 12).first
      at_24 = metrics.text_extents("Hamburgefonstiv", 24).first
      # Integer EMU truncation leaves a little slack.
      expect(at_24).to be_within(2).of(at_12 * 2)
    end

    # A descender and an ascender both have to count, or line height comes out
    # too small and the fitter overpacks the box.
    it "measures ink height, so a descender makes text taller" do
      flat = metrics.text_extents("ace", 18).last
      tall = metrics.text_extents("Ty", 18).last
      expect(tall).to be > flat
    end

    it "gives a space no ink but still advances the pen" do
      aggregate_failures do
        expect(metrics.text_extents(" ", 18)).to eq([0, 0])
        expect(metrics.text_extents("a a", 18).first).to be > metrics.text_extents("aa", 18).first
      end
    end

    # A space has no outline at all, so it must contribute advance and nothing
    # else -- inventing a zero-size box for it would drag the right edge out to
    # the pen position.
    it "lets trailing whitespace advance the pen without adding ink" do
      aggregate_failures do
        expect(metrics.text_extents("a ", 18).first).to eq(metrics.text_extents("a", 18).first)
        expect(metrics.text_extents("a  ", 18).first).to eq(metrics.text_extents("a", 18).first)
      end
    end

    # A monospace font stores one advance for a long run of trailing glyphs, so
    # hmtx is far shorter than the glyph count and every ordinary character is
    # past its end. Reading it without clamping walks off into the next table.
    it "handles a font whose hmtx is shorter than its glyph count" do
      mono = described_class.open(font_or_skip(MONO))
      widths = %w[a b W i].map { |ch| mono.text_extents(ch * 5, 18).first }
      aggregate_failures do
        expect(widths).to all(be > 0)
        # Every advance is identical in a monospace font, so five of anything
        # spans very nearly the same distance.
        expect(widths.max - widths.min).to be < (widths.min * 0.15)
      end
    end

    it "caches by real path so a font is parsed once" do
      expect(described_class.open(SANS)).to be(described_class.open(SANS))
    end

    it "distinguishes two different fonts" do
      font_or_skip(SERIF)
      sans = described_class.open(SANS).text_extents("Hamburgefonstiv", 18).first
      serif = described_class.open(SERIF).text_extents("Hamburgefonstiv", 18).first
      expect(serif).not_to eq(sans)
    end

    # CFF glyph outlines are not parsed; the fallback uses advance widths, so
    # it still has to produce sane, monotonic measurements.
    it "falls back to advance widths for an OpenType/CFF font" do
      cff = described_class.open(font_or_skip(CFF))
      aggregate_failures do
        expect(cff.text_extents("WWWWW", 18).first).to be > cff.text_extents("iiiii", 18).first
        expect(cff.text_extents("Ty", 18).last).to be > 0
      end
    end

    it "refuses a file that is not a font" do
      Tempfile.create(["not", ".ttf"]) do |file|
        file.write("this is not a font")
        file.close
        expect { described_class.new(file.path) }
          .to raise_error(Pptx::Error, /not a TrueType or OpenType font/)
      end
    end
  end

  describe Pptx::TextFitter do
    let(:font) { font_or_skip(SANS) }

    def fitter(text, width_in, height_in)
      described_class.new(text, extents: [Pptx.inches(width_in), Pptx.inches(height_in)],
                                font_file: font)
    end

    it "wraps text onto as many lines as it needs" do
      expect(fitter("one two three four five six", 1.5, 10).wrap(18).size).to be > 1
    end

    it "wraps onto fewer lines at a smaller size" do
      f = fitter("one two three four five six seven eight", 2, 10)
      expect(f.wrap(8).size).to be < f.wrap(24).size
    end

    it "breaks only at word boundaries" do
      lines = fitter("alpha beta gamma delta", 1.2, 10).wrap(14)
      aggregate_failures do
        expect(lines.join(" ")).to eq("alpha beta gamma delta")
        expect(lines).to all(match(/\A\S+( \S+)*\z/))
      end
    end

    # python-pptx raises TypeError here; see PORTING.md. Putting the word on a
    # line of its own lets the height check reject the size instead.
    it "keeps a word too wide for the box on a line of its own" do
      expect(fitter("Supercalifragilisticexpialidocious", 0.5, 10).wrap(18))
        .to eq(["Supercalifragilisticexpialidocious"])
    end

    it "does not hang on a word wider than the box" do
      size = described_class.best_fit_font_size(
        "Supercalifragilisticexpialidocious",
        extents: [Pptx.inches(0.5), Pptx.inches(2)], max_size: 18, font_file: font
      )
      expect(size).to be_between(1, 18)
    end

    it "never exceeds max_size, however much room there is" do
      size = described_class.best_fit_font_size("Hi", extents: [Pptx.inches(40), Pptx.inches(40)],
                                                      max_size: 18, font_file: font)
      expect(size).to eq(18)
    end

    it "returns nil when the text cannot fit at even one point" do
      size = described_class.best_fit_font_size("Some text here", extents: [1, 1],
                                                                  max_size: 18, font_file: font)
      expect(size).to be_nil
    end

    it "picks a smaller size for a smaller box" do
      text = "Annual revenue report"
      big = described_class.best_fit_font_size(text, max_size: 54, font_file: font,
                                                     extents: [Pptx.inches(8), Pptx.inches(3)])
      small = described_class.best_fit_font_size(text, max_size: 54, font_file: font,
                                                       extents: [Pptx.inches(2), Pptx.inches(0.6)])
      expect(small).to be < big
    end

    # The whole point of the search: the size it returns fits and the next one
    # up does not.
    it "returns the largest size that actually fits" do
      text = "Fit this sentence into the available space"
      extents = [Pptx.inches(3), Pptx.inches(1)]
      size = described_class.best_fit_font_size(text, extents: extents, max_size: 54, font_file: font)
      f = described_class.new(text, extents: extents, font_file: font)
      aggregate_failures do
        expect(f.best_fit(size)).to eq(size)
        expect(f.best_fit(size + 1)).to eq(size)
      end
    end
  end

  describe "TextFrame#fit_text" do
    let(:presentation) { Pptx::Presentation.new_default }
    let(:slide) { presentation.slides.add(presentation.slide_layouts[6]) }
    let(:shape) do
      slide.shapes.add_textbox(at: [Pptx.inches(1), Pptx.inches(1)],
                               size: [Pptx.inches(4), Pptx.inches(1.5)])
    end
    let(:frame) { shape.text_frame }

    before { font_or_skip(SANS) }

    it "sets word wrap on and autofit off" do
      frame.text = "Some text to fit"
      frame.fit_text(font_file: SANS)
      aggregate_failures do
        expect(frame.word_wrap).to be(true)
        expect(frame.auto_size).to eq(Pptx::Enum::MSO_AUTO_SIZE::NONE)
      end
    end

    it "applies the fitted size and font to every run" do
      frame.text = "First line\nSecond line"
      size = frame.fit_text(font_file: SANS, font_family: "DejaVu Sans", max_size: 20)
      fonts = frame.paragraphs.flat_map { |p| p.runs.map(&:font) }
      aggregate_failures do
        expect(fonts.size).to eq(2)
        expect(fonts.map(&:name)).to all(eq("DejaVu Sans"))
        expect(fonts.map { |f| f.size.pt }).to all(eq(size))
      end
    end

    it "carries bold and italic through to the runs" do
      frame.text = "Bold and italic"
      frame.fit_text(font_file: SANS, bold: true, italic: true)
      font = frame.paragraphs.first.runs.first.font
      aggregate_failures do
        expect(font.bold).to be(true)
        expect(font.italic).to be(true)
      end
    end

    # So that typing after the last run in PowerPoint continues in the same
    # font rather than reverting to the placeholder default.
    it "sets the end-paragraph properties too" do
      frame.text = "Some text"
      frame.fit_text(font_file: SANS, font_family: "DejaVu Sans")
      end_properties = frame.element.p_list.first.endParaRPr
      aggregate_failures do
        expect(end_properties).not_to be_nil
        expect(Pptx::Font.new(end_properties).name).to eq("DejaVu Sans")
      end
    end

    it "measures the shape less its margins" do
      frame.text = "A sentence long enough to need wrapping in this box"
      roomy = frame.fit_text(font_file: SANS, max_size: 54)

      frame.margin_left = Pptx.inches(1.5)
      frame.margin_right = Pptx.inches(1.5)
      cramped = frame.fit_text(font_file: SANS, max_size: 54)

      expect(cramped).to be < roomy
    end

    it "shrinks text further when the shape is smaller" do
      text = "A reasonably long sentence that has to be made to fit"
      frame.text = text
      big = frame.fit_text(font_file: SANS, max_size: 54)

      small_shape = slide.shapes.add_textbox(at: [0, 0], size: [Pptx.inches(2), Pptx.inches(0.5)])
      small_shape.text_frame.text = text
      small = small_shape.text_frame.fit_text(font_file: SANS, max_size: 54)

      expect(small).to be < big
    end

    it "refuses to fit an empty text frame" do
      expect { frame.fit_text(font_file: SANS) }
        .to raise_error(Pptx::Error, /no text/)
    end

    it "reports when the text cannot be made to fit" do
      tiny = slide.shapes.add_textbox(at: [0, 0], size: [1000, 1000])
      tiny.text_frame.text = "Far too much text for this box"
      expect { tiny.text_frame.fit_text(font_file: SANS, max_size: 18) }
        .to raise_error(Pptx::Error, /does not fit/)
    end
  end

  # The size is measured from glyph outlines rather than by rasterizing, so it
  # is compared with a tolerance; everything fit_text then *writes* is checked
  # exactly, by the package comparison below.
  describe "agreement with python-pptx" do
    before do
      require_oracle!
      font_or_skip(SANS)
    end

    # Pillow ships with python-pptx, so the oracle guard already covers it.
    def pillow_widths(font, rows)
      script = <<~PY
        import json, sys
        from PIL import ImageFont
        out = []
        for text, point_size in json.loads(sys.argv[1]):
            font = ImageFont.truetype(sys.argv[2], point_size)
            left, _, right, _ = font.getbbox(text)
            out.append(right - left)
        json.dump(out, sys.stdout)
      PY
      out, err, status = Open3.capture3("python3", "-c", script, JSON.dump(rows), font)
      raise "pillow oracle failed: #{err}" unless status.success?

      JSON.parse(out)
    end

    def python_sizes(cases)
      script = <<~PY
        import json, sys
        from pptx.text.layout import TextFitter
        out = []
        for text, w, h, max_size in json.loads(sys.argv[1]):
            try:
                out.append(TextFitter.best_fit_font_size(text, (w, h), max_size, #{SANS.inspect}))
            except TypeError:
                # Raised when one word is too wide to fit at the trial size.
                out.append(None)
        json.dump(out, sys.stdout)
      PY
      out, err, status = Open3.capture3("python3", "-c", script, JSON.dump(cases))
      raise "text fitter oracle failed: #{err}" unless status.success?

      JSON.parse(out)
    end

    CORPUS = [
      "Annual Report",
      "Fit this text into the box",
      "The quick brown fox jumps over the lazy dog",
      "Revenue grew 24% year over year across every region",
      "One two three four five six seven eight nine ten",
      "Highlights and outlook for the coming fiscal year",
      "Prepared in Ruby by the reporting team"
    ].freeze

    BOXES = [[4, 1], [6, 2], [3, 0.8], [9, 1.5], [8, 4]].freeze

    def corpus_cases
      CORPUS.product(BOXES, [18, 36, 54]).map do |text, (w, h), max_size|
        [text, Pptx.inches(w).to_i, Pptx.inches(h).to_i, max_size]
      end
    end

    it "chooses a size within one point of python-pptx" do
      cases = corpus_cases
      theirs = python_sizes(cases)
      compared = 0
      mismatches = cases.zip(theirs).filter_map do |(text, w, h, max_size), expected|
        # nil is a case python-pptx cannot answer at all; those are covered by
        # the example below rather than here.
        next if expected.nil?

        compared += 1
        ours = Pptx::TextFitter.best_fit_font_size(text, extents: [w, h], max_size: max_size,
                                                         font_file: SANS)
        next if ours && (ours - expected).abs <= 1

        "#{text.inspect} in #{w}x#{h} max #{max_size}: python #{expected}, ruby #{ours.inspect}"
      end

      aggregate_failures do
        expect(compared).to be > 50
        expect(mismatches).to eq([])
      end
    end

    # A tolerance of one point is only meaningful if the sizes usually agree
    # exactly -- otherwise the measurement could be badly wrong and still pass.
    it "chooses exactly the same size in the large majority of cases" do
      cases = corpus_cases
      theirs = python_sizes(cases)
      comparable = cases.zip(theirs).compact
      exact = comparable.count do |(text, w, h, max_size), expected|
        Pptx::TextFitter.best_fit_font_size(text, extents: [w, h], max_size: max_size,
                                                  font_file: SANS) == expected
      end
      expect(exact.to_f / comparable.size).to be >= 0.8
    end

    # The sizes agreeing is the thing that matters, but it is an indirect
    # check: a measurement can be wrong by a few percent and still round to the
    # same whole point size. This pins the measurement itself against the
    # rasterizer python-pptx uses.
    it "measures text within a few pixels of Pillow" do
      [SANS, MONO].each do |font|
        next unless File.exist?(font)

        rows = ["Annual Report", "The quick brown fox", "Hello", "WWWWW", "iiiii",
                "Revenue grew 24%"].product([12, 18, 36])
        expected = pillow_widths(font, rows)
        errors = rows.each_with_index.map do |(text, point_size), i|
          emu = Pptx::FontMetrics.open(font).text_extents(text, point_size).first
          ((emu / 914_400.0 * 72) - expected[i]).abs
        end
        aggregate_failures(File.basename(font)) do
          expect(errors.max).to be < 5
          expect(errors.sum / errors.size).to be < 2
        end
      end
    end

    # A documented, deliberate divergence: upstream cannot answer these.
    it "answers cases where python-pptx raises TypeError" do
      text = "Supercalifragilisticexpialidocious and more"
      extents = [Pptx.inches(1).to_i, Pptx.inches(3).to_i]
      expect(python_sizes([[text, *extents, 54]])).to eq([nil])
      expect(Pptx::TextFitter.best_fit_font_size(text, extents: extents, max_size: 54,
                                                       font_file: SANS))
        .to be_between(1, 54)
    end

    # The measurement differs, but what fit_text writes once it has a size must
    # not. Feeding python-pptx's answer in as max_size pins both to that size,
    # so the resulting packages have to be byte-identical.
    it "writes the same XML as python-pptx for a given size" do
      text = "Fit this text into the box"
      size = python_sizes([[text, Pptx.inches(4).to_i, Pptx.inches(1.5).to_i, 18]]).first
      raise "oracle could not fit the fixture" if size.nil?

      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1.5))
          box.text_frame.text = #{text.inspect}
          box.text_frame.fit_text(font_family="DejaVu Sans", max_size=#{size},
                                  font_file=#{SANS.inspect})
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          box = slide.shapes.add_textbox(at: [Pptx.inches(1), Pptx.inches(1)],
                                         size: [Pptx.inches(4), Pptx.inches(1.5)])
          box.text_frame.text = text
          box.text_frame.fit_text(font_file: SANS, font_family: "DejaVu Sans", max_size: size)
          prs.save(path)
        }
      )
    end

    # Multiple paragraphs exercise the endParaRPr of each one, not just the last.
    it "writes the same XML for multi-paragraph text" do
      text = "First line here\nSecond line here"
      size = 12
      expect_same_package(
        python: <<~PY,
          from pptx.util import Inches
          prs = pptx.Presentation()
          slide = prs.slides.add_slide(prs.slide_layouts[6])
          box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))
          box.text_frame.text = #{text.inspect}
          box.text_frame.fit_text(font_family="DejaVu Sans", max_size=#{size}, bold=True,
                                  font_file=#{SANS.inspect})
          prs.save(out)
        PY
        ruby: lambda { |path|
          prs = Pptx::Presentation.new_default
          slide = prs.slides.add(prs.slide_layouts[6])
          box = slide.shapes.add_textbox(at: [Pptx.inches(1), Pptx.inches(1)],
                                         size: [Pptx.inches(5), Pptx.inches(2)])
          box.text_frame.text = text
          box.text_frame.fit_text(font_file: SANS, font_family: "DejaVu Sans", max_size: size,
                                  bold: true)
          prs.save(path)
        }
      )
    end
  end
end
