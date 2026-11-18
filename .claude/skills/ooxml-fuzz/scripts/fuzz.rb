# frozen_string_literal: true

# Fuzz ruby_pptx's writer and check every deck with the Open XML SDK's
# OpenXmlValidator (.NET). See ../SKILL.md for the workflow.
#
#   bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb run      [--count 50] [--seed N] [--ops 12] [--edge] [--out DIR]
#   bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb replay   CASE.json [--keep 0,3,5] [--out FILE.pptx]
#   bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb minimize CASE.json [--signature SIG]
#   bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb validate FILE.pptx...
#
# Every deck is described by a list of ops: plain JSON hashes drawn from a
# seeded RNG. A case file (<out>/case-<seed>.json) holds that list, so a
# finding replays exactly, and the minimizer can drop ops one by one until
# only the ones needed to reproduce the validator error are left.

require "json"
require "optparse"
require "fileutils"
require "open3"
require "tmpdir"

SKILL_DIR = File.expand_path("..", __dir__)
REPO = File.expand_path("../../..", SKILL_DIR)
$LOAD_PATH.unshift(File.join(REPO, "lib"))
require "ruby_pptx"

FIXTURES = File.join(REPO, "spec", "fixtures")
IMAGES = Dir[File.join(FIXTURES, "images", "*")].reject { |f| f.end_with?(".svg") }.sort
SVG = File.join(FIXTURES, "images", "vector.svg")
MOVIE = File.join(FIXTURES, "movie.mp4")
OLE_FILE = File.join(FIXTURES, "basic.pptx")
FONT_FILE = ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
             "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans.ttf"].find { |f| File.exist?(f) }

# Errors the library raises on purpose for input it refuses. An op that hits
# one of these is "rejected", which is fine. Anything else is a crash.
REJECTIONS = [ArgumentError, TypeError, RangeError, IndexError, NotImplementedError, Pptx::Error].freeze

E = Pptx::Enum

# ---------------------------------------------------------------------------
# Generator: RNG -> list of op hashes. Pure data; nothing touches a deck here.
# ---------------------------------------------------------------------------
class Generator
  OP_WEIGHTS = {
    "slide" => 3, "autoshape" => 6, "textbox" => 5, "picture" => 3, "table" => 4,
    "chart" => 7, "xy_chart" => 3, "combo_chart" => 2, "connector" => 2, "freeform" => 2,
    "group" => 2, "movie" => 1, "ole" => 1, "notes" => 2, "section" => 1, "master" => 1,
    "table_pages" => 1, "core_props" => 1, "slide_size" => 1, "background" => 2, "fit_text" => 1,
    "builder" => 1
  }.freeze

  CATEGORY_CHARTS = E::XL_CHART_TYPE.map(&:name).reject { |n| n.to_s =~ /XY_SCATTER|BUBBLE/ }.freeze
  XY_CHARTS = E::XL_CHART_TYPE.map(&:name).select { |n| n.to_s =~ /XY_SCATTER/ }.freeze
  BUBBLE_CHARTS = E::XL_CHART_TYPE.map(&:name).select { |n| n.to_s =~ /BUBBLE/ }.freeze
  COMBO_PLOTS = %w[column_clustered column_stacked bar_clustered line line_markers area area_stacked].freeze
  AUTOSHAPES = E::MSO_AUTO_SHAPE_TYPE.map(&:name).freeze
  AUTONUM = %w[arabicPeriod arabicParenR romanUcPeriod romanLcPeriod alphaLcParenR alphaUcPeriod].freeze

  def initialize(seed, ops:, edge:)
    @rng = Random.new(seed)
    @ops = ops
    @edge = edge
  end

  # Round-tripped through JSON so a fresh run builds from exactly what the
  # case file records, and a replay is byte-for-byte the same input.
  def generate
    count = int(1, @ops)
    ops = Array.new(count) { op }
    JSON.parse(JSON.generate(ops, allow_nan: true), allow_nan: true)
  end

  private

  def op
    kind = weighted(OP_WEIGHTS)
    send(:"gen_#{kind}").merge("op" => kind, "slide" => int(0, 7))
  end

  # -- primitives -----------------------------------------------------------

  def int(lo, hi) = @rng.rand(lo..hi)
  def float(lo, hi) = @rng.rand * (hi - lo) + lo
  def chance(p = 0.5) = @rng.rand < p
  def pick(list) = list[@rng.rand(list.size)]
  def some(list, max = 3) = list.sample(int(0, [max, list.size].min), random: @rng)

  def weighted(weights)
    total = weights.values.sum
    r = @rng.rand(total)
    weights.each { |k, w| return k if (r -= w).negative? }
  end

  # Positions and sizes in EMU. Edge mode adds zero, huge and negative.
  def emu
    return pick([0, 1, -914_400, 2**31 - 1, 2**31, 51_206_400 * 2]) if @edge && chance(0.15)

    int(0, 9_144_000)
  end

  def box = { "at" => [emu, emu], "size" => [emu, emu] }

  def rgb
    return pick(["GGGGGG", "12345", "#ABCDEF", "abcdef", ""]) if @edge && chance(0.1)

    format("%06X", @rng.rand(0x1000000))
  end

  def text
    pool = ["Hello", "Quarterly Review", "Ünïcödé ✓ 日本語 עברית", "emoji 🦊🚀",
            "a & b < c > d \" ' ", "", " leading and trailing ", "line one\nline two",
            "tab\there", "soft\vbreak", "x" * int(1, 400), "1234567890", "<a:t>not xml</a:t>"]
    edge_pool = ["ctrl \u0001\u0008\u000B\u001F chars", "lone surrogate-ish \u{FFFE}\u{FFFF}",
                 "nul \u0000 byte", "\r\n windows \r newline", "]]> cdata end", "x" * 40_000]
    pick(@edge && chance(0.3) ? edge_pool : pool)
  end

  def color
    if chance(0.6)
      { "rgb" => rgb, "brightness" => (chance(0.3) ? brightness : nil) }
    else
      { "theme" => pick(E::MSO_THEME_COLOR_INDEX.map(&:name).reject { |n| n == :NOT_THEME_COLOR }).to_s,
        "brightness" => (chance(0.4) ? brightness : nil) }
    end
  end

  def brightness = @edge && chance(0.1) ? pick([-1.5, 2, Float::NAN]) : float(-1, 1).round(3)

  def fill
    case int(0, 5)
    when 0 then { "type" => "solid", "color" => color }
    when 1
      { "type" => "gradient", "angle" => (@edge && chance(0.2) ? pick([-90, 720, 359.999]) : int(0, 359)),
        "stops" => Array.new(int(0, 3)) { { "color" => color, "position" => float(0, 1).round(3) } } }
    when 2
      { "type" => "patterned", "pattern" => pick(E::MSO_PATTERN_TYPE.map(&:name)).to_s,
        "fore" => (chance(0.8) ? color : nil), "back" => (chance(0.6) ? color : nil) }
    when 3 then { "type" => "background" }
    else nil
    end
  end

  def line
    return nil if chance(0.3)

    { "color" => (chance(0.7) ? color : nil),
      "width" => (chance(0.7) ? int(0, @edge ? 20_116_800 : 254_000) : nil),
      "dash" => (chance(0.4) ? pick(E::MSO_LINE_DASH_STYLE.map(&:name)).to_s : nil),
      "fill" => (chance(0.2) ? fill : nil) }
  end

  def font
    {
      "bold" => (chance(0.3) ? pick([true, false, nil]) : :skip),
      "italic" => (chance(0.3) ? pick([true, false, nil]) : :skip),
      "size" => (chance(0.4) ? (@edge && chance(0.2) ? pick([0, 0.5, 4001, 3999.5]) : int(6, 96)) : :skip),
      "name" => (chance(0.3) ? pick(["Arial", "Georgia", "Noto Sans CJK JP", "", "Very \"Odd\" <Font>"]) : :skip),
      "caps" => (chance(0.2) ? pick(%i[all small none]).to_s : :skip),
      "underline" => (chance(0.3) ? pick([true, false, *E::MSO_TEXT_UNDERLINE_TYPE.map { |m| m.name.to_s }]) : :skip),
      "language" => (chance(0.2) ? pick(%w[EN_US DE_CH JAPANESE ARABIC HEBREW]) : :skip),
      "color" => (chance(0.4) ? color : :skip)
    }.reject { |_, v| v == :skip }
  end

  def paragraphs
    Array.new(int(1, 4)) do
      {
        "runs" => Array.new(int(0, 3)) do
          { "text" => text, "font" => font, "hyperlink" => (chance(0.1) ? url : nil),
            "break_after" => chance(0.1) }
        end,
        "field" => (chance(0.1) ? pick(%w[slide_number date_time datetime4]) : nil),
        "alignment" => (chance(0.4) ? pick(E::PP_PARAGRAPH_ALIGNMENT.map(&:name).reject { |n| n == :MIXED }).to_s : nil),
        "level" => (chance(0.3) ? int(0, @edge ? 12 : 8) : nil),
        "line_spacing" => (chance(0.3) ? (chance ? float(0.5, 3).round(2) : { "pt" => int(1, 200) }) : nil),
        "space_before" => (chance(0.2) ? int(0, 1_000) : nil),
        "space_after" => (chance(0.2) ? int(0, 1_000) : nil),
        "bullet" => (chance(0.3) ? pick([nil, "none", "•", "–", { "autonum" => pick(@edge && chance(0.2) ? %w[bogus arabicperiod] : AUTONUM) }]) : :skip),
        "indent" => (chance(0.15) ? int(-914_400, 914_400) : nil),
        "margin_left" => (chance(0.15) ? int(0, 914_400) : nil),
        "font" => (chance(0.2) ? font : nil)
      }.reject { |_, v| v == :skip }
    end
  end

  def text_frame
    {
      "paragraphs" => paragraphs,
      "margins" => (chance(0.3) ? Array.new(4) { int(0, 457_200) } : nil),
      "anchor" => (chance(0.3) ? pick(%w[TOP MIDDLE BOTTOM]) : nil),
      "auto_size" => (chance(0.3) ? pick(%w[NONE SHAPE_TO_FIT_TEXT TEXT_TO_FIT_SHAPE]) : nil),
      "word_wrap" => (chance(0.3) ? pick([true, false, nil]) : :skip),
      "direction" => (chance(0.15) ? pick(%w[horz vert vert270 wordArtVert eaVert mongolianVert wordArtVertRtl]) : nil)
    }.reject { |_, v| v == :skip }
  end

  def url = pick(["https://example.com", "https://example.com/a?b=1&c=<2>", "mailto:x@example.com",
                  "file:///C:/path with spaces/x.pptx", "#local", "ftp://ü.example/ä"])

  def action
    return nil unless chance(0.25)

    chance ? { "hyperlink" => url } : { "target_slide" => int(0, 7) }
  end

  # -- ops -----------------------------------------------------------------

  def gen_slide = { "layout" => int(0, 10), "title" => (chance(0.7) ? text : nil), "body" => (chance(0.5) ? text_frame : nil) }

  def gen_autoshape
    {
      "type" => pick(AUTOSHAPES).to_s, **box, "fill" => fill, "line" => line,
      "rotation" => (chance(0.3) ? (@edge ? pick([-720.5, 0, 359.99, 1e9]) : float(0, 360).round(2)) : nil),
      "adjustments" => (chance(0.3) ? Array.new(int(1, 4)) { float(-1, 2).round(4) } : nil),
      "text" => (chance(0.6) ? text_frame : nil), "shadow_inherit" => (chance(0.2) ? chance : nil),
      "action" => action, "name" => (chance(0.1) ? text : nil)
    }
  end

  def gen_textbox = { **box, "text" => text_frame, "fill" => (chance(0.3) ? fill : nil), "action" => action }

  def gen_picture
    svg = chance(0.15)
    {
      "image" => File.basename(svg ? SVG : pick(IMAGES)), "svg" => svg, "fallback" => File.basename(pick(IMAGES)),
      "at" => [emu, emu], "width" => (chance ? emu : nil), "height" => (chance ? emu : nil),
      "crop" => (chance(0.3) ? Array.new(4) { float(@edge ? -0.5 : 0, @edge ? 1.5 : 0.4).round(3) } : nil),
      "line" => line, "rotation" => (chance(0.2) ? float(0, 360).round(1) : nil), "action" => action
    }
  end

  def gen_table
    rows = int(1, @edge ? 30 : 8)
    cols = int(1, @edge ? 30 : 6)
    {
      "rows" => rows, "cols" => cols, **box,
      "flags" => %w[first_row first_col last_row last_col banded_rows banded_columns].to_h { |f| [f, chance ? chance : nil] },
      "cells" => Array.new(int(0, rows * cols)) do
        { "r" => int(0, rows - 1), "c" => int(0, cols - 1), "text" => (chance(0.8) ? text : nil),
          "fill" => (chance(0.3) ? fill : nil), "anchor" => (chance(0.2) ? pick(%w[TOP MIDDLE BOTTOM]) : nil),
          "margins" => (chance(0.2) ? Array.new(4) { int(0, 228_600) } : nil) }
      end,
      "merges" => Array.new(int(0, 2)) { [int(0, rows - 1), int(0, cols - 1), int(0, rows - 1), int(0, cols - 1)] },
      "col_widths" => (chance(0.3) ? Array.new(cols) { int(@edge ? 0 : 1, 3_000_000) } : nil),
      "row_heights" => (chance(0.3) ? Array.new(rows) { int(@edge ? 0 : 1, 1_000_000) } : nil)
    }
  end

  def number
    return pick([nil, 0, -1e308, 1e308, Float::NAN, Float::INFINITY, -0.0, 1e-300]) if @edge && chance(0.15)
    return nil if chance(0.05)

    chance(0.7) ? float(-100, 1000).round(int(0, 4)) : int(-1000, 100_000)
  end

  def categories
    n = int(@edge ? 0 : 1, 8)
    case int(0, 3)
    when 0 then Array.new(n) { |i| "Cat #{i} #{chance(0.2) ? text[0, 20] : ''}" }
    when 1 then Array.new(n) { |i| i * int(1, 10) }
    when 2 then Array.new(n) { |i| { "date" => "2026-#{format('%02d', (i % 12) + 1)}-01" } }
    else { "levels" => [Array.new(n) { |i| "Leaf #{i}" }, Array.new([n, 1].max) { |i| "Group #{i / 2}" }] }
    end
  end

  def chart_data
    cats = categories
    count = cats.is_a?(Hash) ? cats["levels"].first.size : cats.size
    {
      "categories" => cats,
      "number_format" => (chance(0.2) ? pick(['0.0%', '#,##0', '"$"#,##0.00', 'General', 'yyyy-mm-dd']) : nil),
      "series" => Array.new(int(1, 5)) do |i|
        { "name" => chance(0.8) ? "Series #{i}" : text, "values" => Array.new(count) { number } }
      end
    }
  end

  def chart_dressing
    {
      "title" => (chance(0.4) ? pick([nil, text, true]) : :skip),
      "legend" => (chance(0.5) ? { "on" => chance(0.8), "position" => pick(%w[BOTTOM CORNER LEFT RIGHT TOP]),
                                   "in_layout" => chance, "font" => font } : nil),
      "style" => (chance(0.2) ? int(@edge ? 0 : 1, @edge ? 99 : 48) : nil),
      "font" => (chance(0.2) ? font : nil),
      "value_axis" => (chance(0.4) ? axis : nil), "category_axis" => (chance(0.3) ? axis : nil),
      "plot" => (chance(0.5) ? plot : nil),
      "series_fmt" => (chance(0.5) ? series_fmt : nil),
      "replace" => (chance(0.15) ? chart_data : nil)
    }.reject { |_, v| v == :skip }
  end

  def axis
    {
      "major_gridlines" => (chance ? chance : nil), "minor_gridlines" => (chance(0.3) ? chance : nil),
      "major_tick" => (chance(0.3) ? pick(%w[CROSS INSIDE NONE OUTSIDE]) : nil),
      "minor_tick" => (chance(0.2) ? pick(%w[CROSS INSIDE NONE OUTSIDE]) : nil),
      "tick_label_position" => (chance(0.2) ? pick(%w[HIGH LOW NEXT_TO_AXIS NONE]) : nil),
      "max" => (chance(0.2) ? number : :skip), "min" => (chance(0.2) ? number : :skip),
      "major_unit" => (chance(0.15) ? number : :skip),
      "reverse" => (chance(0.2) ? chance : nil), "visible" => (chance(0.15) ? chance : nil),
      "number_format" => (chance(0.2) ? pick(['0%', '#,##0.0', '@']) : nil),
      "title" => (chance(0.3) ? text : nil), "crosses" => (chance(0.15) ? pick(%w[AUTOMATIC MAXIMUM MINIMUM]) : nil)
    }.reject { |_, v| v == :skip }
  end

  def plot
    {
      "gap_width" => (chance ? int(@edge ? -10 : 0, @edge ? 600 : 500) : nil),
      "overlap" => (chance ? int(@edge ? -200 : -100, @edge ? 200 : 100) : nil),
      "vary" => (chance(0.3) ? chance : nil),
      "labels" => (chance(0.6) ? {
        "show" => %w[value category_name series_name percentage legend_key bubble_size].select { chance(0.3) },
        "position" => (chance(0.4) ? pick(E::XL_DATA_LABEL_POSITION.map(&:name)).to_s : nil),
        "number_format" => (chance(0.3) ? pick(['0.0', '0%']) : nil), "font" => (chance(0.3) ? font : nil)
      } : nil)
    }
  end

  def series_fmt
    {
      "index" => int(0, 4), "fill" => fill, "line" => line,
      "marker" => (chance ? { "style" => pick(E::XL_MARKER_STYLE.map(&:name)).to_s,
                              "size" => int(@edge ? 0 : 2, @edge ? 100 : 72) } : nil),
      "smooth" => (chance(0.3) ? chance : nil), "invert" => (chance(0.2) ? chance : nil),
      "point" => (chance(0.3) ? { "index" => int(0, 7), "fill" => fill, "label" => text } : nil)
    }
  end

  def gen_chart = { "type" => pick(CATEGORY_CHARTS).to_s, **box, "data" => chart_data, **chart_dressing }

  def gen_xy_chart
    bubble = chance(0.4)
    {
      "type" => pick(bubble ? BUBBLE_CHARTS : XY_CHARTS).to_s, **box, "bubble" => bubble,
      "series" => Array.new(int(1, 4)) do |i|
        { "name" => "S#{i}", "points" => Array.new(int(@edge ? 0 : 1, 10)) { bubble ? [number, number, number] : [number, number] } }
      end,
      **chart_dressing.except("replace")
    }
  end

  def gen_combo_chart
    data = chart_data
    data["series"] << { "name" => "Extra", "values" => Array.new(data["series"].first["values"].size) { number } } if data["series"].size < 2
    {
      **box, "data" => data,
      "plots" => data["series"].each_index.map { |i| { "type" => pick(COMBO_PLOTS), "series" => i, "secondary" => i.positive? && chance } },
      **chart_dressing.except("replace")
    }
  end

  def gen_connector
    { "type" => pick(%w[STRAIGHT ELBOW CURVE]), "begin" => [emu, emu], "end" => [emu, emu],
      "connect" => (chance ? [int(0, 20), int(0, @edge ? 40 : 3)] : nil), "line" => line }
  end

  def gen_freeform
    { "start" => [int(0, 1000), int(0, 1000)], "scale" => (chance ? float(0.1, 10000).round(2) : 1.0),
      "vertices" => Array.new(int(@edge ? 0 : 1, 30)) { [int(-500, 5000), int(-500, 5000)] },
      "close" => chance, "fill" => fill, "line" => line }
  end

  def gen_group = { "shapes" => Array.new(int(@edge ? 0 : 1, 4)) { gen_autoshape.slice("type", "at", "size", "fill") }, "nested" => chance(0.3) }
  def gen_movie = { **box, "poster" => (chance ? File.basename(pick(IMAGES)) : nil) }
  def gen_ole = { "at" => [emu, emu], "icon" => (chance(0.3) ? File.basename(pick(IMAGES)) : nil) }
  def gen_notes = { "text" => text, "bold_last" => chance(0.3) }
  def gen_section = { "name" => text, "count" => int(0, 3) }

  def gen_master
    {
      "name" => text, "placeholders" => pick(%w[standard none]),
      "colors" => (chance ? %w[accent1 accent2 dk1 lt1 hlink].to_h { |k| [k, rgb] }.select { chance } : nil),
      "major" => (chance ? pick(["Georgia", "", "Ünï"]) : nil), "minor" => (chance ? "Verdana" : nil),
      "layout_type" => pick(%w[obj title twoObj blank txAndChart picTx tbl chart]),
      "title" => text
    }
  end

  def gen_table_pages = { "rows" => Array.new(int(0, 60)) { Array.new(int(1, 4)) { text[0, 30] } }, "header" => chance(0.7), **box }

  def gen_core_props
    %w[title author subject keywords comments category last_modified_by content_status
       identifier language version].to_h { |k| [k, chance(0.4) ? text : nil] }.compact
      .merge("revision" => (chance(0.3) ? int(@edge ? -1 : 1, 999_999) : nil))
  end

  def gen_slide_size = { "width" => int(@edge ? 0 : 914_400, @edge ? 2**31 : 51_206_400), "height" => int(@edge ? 0 : 914_400, 51_206_400) }
  def gen_background = { "fill" => fill || { "type" => "solid", "color" => color } }
  def gen_fit_text = { **box, "text" => text, "max" => int(@edge ? 0 : 6, 72) }
  def gen_builder = { "title" => text, "shape" => pick(AUTOSHAPES).to_s, "fill" => rgb, "text" => text }
end

# ---------------------------------------------------------------------------
# Executor: op hashes -> a saved deck.
# ---------------------------------------------------------------------------
class Executor
  attr_reader :log

  def initialize(ops)
    @ops = ops
    @log = [] # one entry per op: "ok", "rejected: ...", "crash: ..."
  end

  def build
    @prs = Pptx::Presentation.new_default
    @ops.each_with_index do |op, i|
      begin
        send(:"op_#{op['op']}", op)
        @log << { "i" => i, "op" => op["op"], "status" => "ok", "notes" => @notes.to_a }
      rescue *REJECTIONS => e
        @log << { "i" => i, "op" => op["op"], "status" => "rejected", "error" => "#{e.class}: #{e.message[0, 300]}" }
      rescue StandardError, SystemStackError => e
        @log << { "i" => i, "op" => op["op"], "status" => "crash", "error" => "#{e.class}: #{e.message[0, 300]}",
                  "backtrace" => e.backtrace.first(8), "in_library" => in_library?(e) }
      end
      @notes = nil
    end
    @prs
  end

  def save(path)
    @prs.save(path)
  end

  private

  def in_library?(error) = Executor.in_library?(error)

  # Whether the exception came out of ruby_pptx (or a gem it called), rather
  # than from this harness calling something that does not exist.
  def self.in_library?(error)
    frame = error.backtrace.to_a.find { |f| !f.include?("/gems/") && !f.include?("<internal") }
    frame.to_s.start_with?(File.join(REPO, "lib")) || frame.to_s.start_with?("lib/ruby_pptx")
  end

  # Run one optional mutation. A rejection here is noted but does not abort
  # the rest of the op, so one refused value does not hide later ones.
  def try(label)
    yield
  rescue *REJECTIONS => e
    (@notes ||= []) << "#{label} rejected: #{e.class}: #{e.message[0, 160]}"
  rescue NoMethodError => e
    # The API this op reached for does not exist on this kind of object (no
    # markers on a bar series, no major_unit on a category axis). Only a
    # NoMethodError raised from inside the library is a crash.
    raise if in_library?(e)

    (@notes ||= []) << "#{label} n/a: #{e.message[0, 160]}"
  end

  def len(v) = Pptx.emu(Integer(v))
  def pos(xy) = [len(xy[0]), len(xy[1])]

  def slide_for(op)
    @prs.slides.add(@prs.slide_layouts[6] || @prs.slide_layouts[0]) if @prs.slides.size.zero?
    @prs.slides[op["slide"] % @prs.slides.size]
  end

  def shapes(op) = slide_for(op).shapes

  def file(name) = [*IMAGES, SVG].find { |f| File.basename(f) == name }

  # -- dml ------------------------------------------------------------------

  def apply_color(color_format, spec)
    return if spec.nil?

    if spec["rgb"]
      try("rgb") { color_format.rgb = spec["rgb"] }
    else
      try("theme_color") { color_format.theme_color = spec["theme"].to_sym }
    end
    try("brightness") { color_format.brightness = spec["brightness"] } unless spec["brightness"].nil?
  end

  def apply_fill(fill, spec)
    return if spec.nil?

    case spec["type"]
    when "solid"
      fill.solid
      apply_color(fill.fore_color, spec["color"])
    when "gradient"
      fill.gradient
      try("gradient_angle") { fill.gradient_angle = spec["angle"] }
      spec["stops"].each_with_index do |stop, i|
        s = fill.gradient_stops.to_a[i] or break
        apply_color(s.color, stop["color"])
        try("stop position") { s.position = stop["position"] }
      end
    when "patterned"
      fill.patterned
      try("pattern") { fill.pattern = spec["pattern"].to_sym }
      apply_color(fill.fore_color, spec["fore"])
      apply_color(fill.back_color, spec["back"])
    when "background" then fill.background
    end
  end

  def apply_line(line, spec)
    return if spec.nil?

    apply_fill(line.fill, spec["fill"])
    if spec["color"]
      try("line color") do
        line.fill.solid unless spec["fill"]
        apply_color(line.color, spec["color"])
      end
    end
    try("line width") { line.width = len(spec["width"]) } if spec["width"]
    try("dash") { line.dash_style = spec["dash"].to_sym } if spec["dash"]
  end

  def apply_font(font, spec)
    return if spec.nil?

    spec.each do |k, v|
      try("font.#{k}") do
        case k
        when "size" then font.size = Pptx.pt(v)
        when "color" then apply_color(font.color, v)
        when "caps", "language" then font.public_send(:"#{k}=", v.to_sym)
        when "underline" then font.underline = v.is_a?(String) ? v.to_sym : v
        else font.public_send(:"#{k}=", v)
        end
      end
    end
  end

  def apply_action(shape, spec)
    return if spec.nil?

    if spec["hyperlink"]
      try("click hyperlink") { shape.hyperlink = spec["hyperlink"] }
    else
      try("target_slide") { shape.click_action.target_slide = @prs.slides[spec["target_slide"] % @prs.slides.size] }
    end
  end

  def apply_text_frame(tf, spec)
    return if spec.nil?

    tf.clear
    spec["paragraphs"].each_with_index do |ps, pi|
      para = pi.zero? ? tf.paragraphs.first : tf.add_paragraph
      ps["runs"].each do |rs|
        run = para.add_run
        try("run.text") { run.text = rs["text"] }
        apply_font(run.font, rs["font"])
        try("run.hyperlink") { run.hyperlink = rs["hyperlink"] } if rs["hyperlink"]
        para.add_line_break if rs["break_after"]
      end
      try("field") { para.add_field(ps["field"].match?(/\A[a-z_]+\z/) && ps["field"].include?("_") ? ps["field"].to_sym : ps["field"]) } if ps["field"]
      try("alignment") { para.alignment = ps["alignment"].to_sym } if ps["alignment"]
      try("level") { para.level = ps["level"] } if ps["level"]
      if ps["line_spacing"]
        try("line_spacing") { para.line_spacing = ps["line_spacing"].is_a?(Hash) ? Pptx.pt(ps["line_spacing"]["pt"]) : ps["line_spacing"] }
      end
      try("space_before") { para.space_before = Pptx.pt(ps["space_before"]) } if ps["space_before"]
      try("space_after") { para.space_after = Pptx.pt(ps["space_after"]) } if ps["space_after"]
      if ps.key?("bullet")
        b = ps["bullet"]
        try("bullet") { para.bullet = b.is_a?(Hash) ? b["autonum"].to_sym : (b == "none" ? :none : b) }
      end
      try("indent") { para.indent = len(ps["indent"]) } if ps["indent"]
      try("p.margin_left") { para.margin_left = len(ps["margin_left"]) } if ps["margin_left"]
      apply_font(para.font, ps["font"])
    end
    if spec["margins"]
      %i[margin_left margin_top margin_right margin_bottom].zip(spec["margins"]) { |m, v| try(m) { tf.public_send(:"#{m}=", len(v)) } }
    end
    try("anchor") { tf.vertical_anchor = spec["anchor"].to_sym } if spec["anchor"]
    try("auto_size") { tf.auto_size = spec["auto_size"].to_sym } if spec["auto_size"]
    try("word_wrap") { tf.word_wrap = spec["word_wrap"] } if spec.key?("word_wrap")
    try("text_direction") { tf.text_direction = spec["direction"].to_sym } if spec["direction"]
  end

  # -- ops ------------------------------------------------------------------

  def op_slide(op)
    layouts = @prs.slide_layouts.to_a
    slide = @prs.slides.add(layouts[op["layout"] % layouts.size])
    try("title") { slide.shapes.title.text = op["title"] if slide.shapes.title } if op["title"]
    body = slide.placeholders.to_a.find { |ph| ph.placeholder_format.idx == 1 }
    apply_text_frame(body.text_frame, op["body"]) if body && op["body"] && body.text_frame?
  end

  def op_autoshape(op)
    shape = shapes(op).add_shape(op["type"].to_sym, at: pos(op["at"]), size: pos(op["size"]))
    apply_fill(shape.fill, op["fill"])
    apply_line(shape.line, op["line"])
    try("rotation") { shape.rotation = op["rotation"] } if op["rotation"]
    op["adjustments"]&.each_with_index { |v, i| try("adjustment #{i}") { shape.adjustments[i] = v } }
    apply_text_frame(shape.text_frame, op["text"])
    try("shadow") { shape.shadow.inherit = op["shadow_inherit"] } unless op["shadow_inherit"].nil?
    apply_action(shape, op["action"])
    try("name") { shape.name = op["name"] } if op["name"]
  end

  def op_textbox(op)
    box = shapes(op).add_textbox(at: pos(op["at"]), size: pos(op["size"]))
    apply_text_frame(box.text_frame, op["text"])
    apply_fill(box.fill, op["fill"])
    apply_action(box, op["action"])
  end

  def op_picture(op)
    args = { at: pos(op["at"]) }
    args[:width] = len(op["width"]) if op["width"]
    args[:height] = len(op["height"]) if op["height"]
    args[:fallback] = file(op["fallback"]) if op["svg"]
    pic = shapes(op).add_picture(file(op["image"]), **args)
    op["crop"]&.then { |c| %i[left top right bottom].zip(c) { |s, v| try("crop_#{s}") { pic.public_send(:"crop_#{s}=", v) } } }
    apply_line(pic.line, op["line"])
    try("rotation") { pic.rotation = op["rotation"] } if op["rotation"]
    apply_action(pic, op["action"])
  end

  def op_table(op)
    frame = shapes(op).add_table(op["rows"], op["cols"], at: pos(op["at"]), size: pos(op["size"]))
    table = frame.table
    op["flags"].each { |k, v| try(k) { table.public_send(:"#{k}=", v) } unless v.nil? }
    op["cells"].each do |c|
      cell = table.cell(c["r"], c["c"])
      try("cell.text") { cell.text = c["text"] } if c["text"]
      apply_fill(cell.fill, c["fill"])
      try("cell.anchor") { cell.vertical_anchor = c["anchor"].to_sym } if c["anchor"]
      c["margins"]&.then { |m| %i[margin_left margin_top margin_right margin_bottom].zip(m) { |s, v| try(s) { cell.public_send(:"#{s}=", len(v)) } } }
    end
    op["merges"].each do |r1, c1, r2, c2|
      try("merge") { table.cell(r1, c1).merge(table.cell(r2, c2)) }
    end
    op["col_widths"]&.each_with_index { |w, i| try("col width") { table.columns[i].width = len(w) } }
    op["row_heights"]&.each_with_index { |h, i| try("row height") { table.rows[i].height = len(h) } }
  end

  def build_chart_data(spec)
    data = Pptx::ChartData.new(**(spec["number_format"] ? { number_format: spec["number_format"] } : {}))
    cats = spec["categories"]
    if cats.is_a?(Hash)
      leaves, groups = cats["levels"]
      leaves.each_with_index do |leaf, i|
        parent = data.categories.find { |c| c.label == groups[i / 2] } || data.categories.add_category(groups[i / 2])
        parent.add_sub_category(leaf)
      end
    else
      data.categories = cats.map { |c| c.is_a?(Hash) ? Date.parse(c["date"]) : c }
    end
    spec["series"].each { |s| data.add_series(s["name"], s["values"]) }
    data
  end

  def dress_chart(chart, op)
    try("title") { chart.title = op["title"] } if op.key?("title")
    if (lg = op["legend"])
      try("legend") { chart.legend = lg["on"] }
      if lg["on"]
        try("legend.position") { chart.legend.position = lg["position"].to_sym }
        try("legend.in_layout") { chart.legend.include_in_layout = lg["in_layout"] }
        apply_font(chart.legend.font, lg["font"])
      end
    end
    try("chart_style") { chart.chart_style = op["style"] } if op["style"]
    apply_font(chart.font, op["font"]) if op["font"]
    { "value_axis" => :value_axis, "category_axis" => :category_axis }.each do |key, reader|
      next unless (spec = op[key])

      try(key) do
        axis = chart.public_send(reader)
        try("major_gridlines") { axis.major_gridlines = spec["major_gridlines"] } unless spec["major_gridlines"].nil?
        try("minor_gridlines") { axis.minor_gridlines = spec["minor_gridlines"] } unless spec["minor_gridlines"].nil?
        try("major_tick") { axis.major_tick_mark = spec["major_tick"].to_sym } if spec["major_tick"]
        try("minor_tick") { axis.minor_tick_mark = spec["minor_tick"].to_sym } if spec["minor_tick"]
        try("tick_label_position") { axis.tick_label_position = spec["tick_label_position"].to_sym } if spec["tick_label_position"]
        try("max") { axis.maximum_scale = spec["max"] } if spec.key?("max")
        try("min") { axis.minimum_scale = spec["min"] } if spec.key?("min")
        try("major_unit") { axis.major_unit = spec["major_unit"] } if spec.key?("major_unit")
        try("reverse") { axis.reverse_order = spec["reverse"] } unless spec["reverse"].nil?
        try("visible") { axis.visible = spec["visible"] } unless spec["visible"].nil?
        try("axis number_format") { axis.tick_labels.number_format = spec["number_format"] } if spec["number_format"]
        try("axis title") { axis.title = spec["title"] } if spec["title"]
        try("crosses") { axis.crosses = spec["crosses"].to_sym } if spec["crosses"]
      end
    end
    if (pl = op["plot"])
      try("plot") do
        plot = chart.plots.first
        try("gap_width") { plot.gap_width = pl["gap_width"] } if pl["gap_width"]
        try("overlap") { plot.overlap = pl["overlap"] } if pl["overlap"]
        try("vary") { plot.vary_by_categories = pl["vary"] } unless pl["vary"].nil?
        if (lb = pl["labels"])
          try("data_labels") do
            plot.data_labels = true
            labels = plot.data_labels
            lb["show"].each { |s| try("show_#{s}") { labels.public_send(:"show_#{s}=", true) } }
            try("label position") { labels.position = lb["position"].to_sym } if lb["position"]
            try("label number_format") { labels.number_format = lb["number_format"] } if lb["number_format"]
            apply_font(labels.font, lb["font"])
          end
        end
      end
    end
    if (sf = op["series_fmt"])
      try("series") do
        series = chart.series.to_a[sf["index"] % chart.series.to_a.size]
        apply_fill(series.format.fill, sf["fill"])
        apply_line(series.format.line, sf["line"])
        if sf["marker"]
          try("marker") do
            series.marker.style = sf["marker"]["style"].to_sym
            series.marker.size = sf["marker"]["size"]
          end
        end
        try("smooth") { series.smooth = sf["smooth"] } unless sf["smooth"].nil?
        try("invert") { series.invert_if_negative = sf["invert"] } unless sf["invert"].nil?
        if (pt = sf["point"])
          try("point") do
            point = series.points[pt["index"] % [series.points.size, 1].max] or next
            apply_fill(point.format.fill, pt["fill"])
            point.data_label.text_frame.text = pt["label"]
          end
        end
      end
    end
    try("replace_data") { chart.replace_data(build_chart_data(op["replace"])) } if op["replace"]
  end

  def op_chart(op)
    frame = shapes(op).add_chart(op["type"].to_sym, build_chart_data(op["data"]), at: pos(op["at"]), size: pos(op["size"]))
    dress_chart(frame.chart, op)
  end

  def op_xy_chart(op)
    data = op["bubble"] ? Pptx::BubbleChartData.new : Pptx::XyChartData.new
    op["series"].each do |s|
      data.add_series(s["name"], points: s["points"])
    end
    frame = shapes(op).add_chart(op["type"].to_sym, data, at: pos(op["at"]), size: pos(op["size"]))
    dress_chart(frame.chart, op)
  end

  def op_combo_chart(op)
    data = build_chart_data(op["data"])
    frame = shapes(op).add_combo_chart(data, at: pos(op["at"]), size: pos(op["size"])) do |combo|
      op["plots"].each { |p| combo.plot(p["type"].to_sym, series: p["series"], secondary_axis: p["secondary"]) }
    end
    dress_chart(frame.chart, op)
  end

  def op_connector(op)
    shp = shapes(op)
    conn = shp.add_connector(op["type"].to_sym, begin_at: pos(op["begin"]), end_at: pos(op["end"]))
    if op["connect"]
      targets = shp.to_a.reject { |s| s.equal?(conn) || s.shape_type.to_s.include?("CONNECTOR") }
      unless targets.empty?
        target = targets[op["connect"][0] % targets.size]
        try("begin_connect") { conn.begin_connect(target, op["connect"][1]) }
        try("end_connect") { conn.end_connect(targets[(op["connect"][0] + 1) % targets.size], op["connect"][1]) }
      end
    end
    apply_line(conn.line, op["line"])
  end

  def op_freeform(op)
    builder = shapes(op).build_freeform(start_x: op["start"][0], start_y: op["start"][1], scale: op["scale"])
    builder.add_line_segments(op["vertices"], close: op["close"])
    shape = builder.convert_to_shape
    apply_fill(shape.fill, op["fill"])
    apply_line(shape.line, op["line"])
  end

  def op_group(op)
    shp = shapes(op)
    members = op["shapes"].map do |s|
      shape = shp.add_shape(s["type"].to_sym, at: pos(s["at"]), size: pos(s["size"]))
      apply_fill(shape.fill, s["fill"])
      shape
    end
    group = shp.add_group_shape(members)
    return unless op["nested"]

    inner = group.shapes.add_shape(:oval, at: [Pptx.emu(0), Pptx.emu(0)], size: [Pptx.inches(1), Pptx.inches(1)])
    group.shapes.add_group_shape([inner])
  end

  def op_movie(op)
    args = { at: pos(op["at"]), size: pos(op["size"]), mime_type: "video/mp4" }
    args[:poster_frame] = file(op["poster"]) if op["poster"]
    shapes(op).add_movie(MOVIE, **args)
  rescue ArgumentError => e
    raise unless e.message.include?("mime_type")

    args.delete(:mime_type)
    shapes(op).add_movie(MOVIE, **args, content_type: "video/mp4")
  end

  def op_ole(op)
    args = { prog_id: :PPTX, at: pos(op["at"]) }
    args[:icon_file] = file(op["icon"]) if op["icon"]
    shapes(op).add_ole_object(OLE_FILE, **args)
  end

  def op_notes(op)
    slide = slide_for(op)
    slide.notes = op["text"]
    return unless op["bold_last"]

    run = slide.notes_slide.notes_text_frame.paragraphs.last.runs.first
    run.font.bold = true if run
  end

  def op_section(op)
    slides = @prs.slides.to_a
    @prs.sections.add(op["name"], slides: slides.last([op["count"], slides.size].min))
  end

  def op_master(op)
    master = @prs.slide_masters.add(name: op["name"], placeholders: op["placeholders"].to_sym)
    try("theme colors") { master.theme.colors.update(**op["colors"].transform_keys(&:to_sym)) } if op["colors"]
    try("major font") { master.theme.fonts.major = op["major"] } if op["major"]
    try("minor font") { master.theme.fonts.minor = op["minor"] } if op["minor"]
    layout = master.slide_layouts.add("Fuzz #{op['layout_type']}", type: op["layout_type"]) do |l|
      l.placeholders.add(:title)
      l.placeholders.add(:body, idx: 1, at: [Pptx.inches(0.5), Pptx.inches(1.75)], size: [Pptx.inches(9), Pptx.inches(4)])
    end
    slide = @prs.slides.add(layout)
    try("master slide title") { slide.shapes.title.text = op["title"] if slide.shapes.title }
  end

  def op_table_pages(op)
    layouts = @prs.slide_layouts.to_a
    @prs.slides.add_table_pages(op["rows"], layout: layouts[6] || layouts.last, header: op["header"],
                                left: len(op["at"][0]), top: len(op["at"][1]),
                                width: len(op["size"][0]), height: len(op["size"][1]))
  end

  def op_core_props(op)
    props = @prs.core_properties
    op.except("op", "slide").each { |k, v| try("core.#{k}") { props.public_send(:"#{k}=", v) } unless v.nil? }
  end

  def op_slide_size(op)
    try("slide_width") { @prs.slide_width = len(op["width"]) }
    try("slide_height") { @prs.slide_height = len(op["height"]) }
  end

  def op_background(op)
    apply_fill(slide_for(op).background.fill, op["fill"])
  end

  def op_fit_text(op)
    raise NotImplementedError, "no DejaVuSans.ttf on this machine" unless FONT_FILE

    box = shapes(op).add_textbox(at: pos(op["at"]), size: pos(op["size"]))
    box.text_frame.text = op["text"]
    box.text_frame.fit_text(font_file: FONT_FILE, font_family: "DejaVu Sans", max_size: op["max"])
  end

  # The declarative builder is a separate deck; save it beside the main one
  # by merging its first slide's shapes isn't possible, so build it inside the
  # same presentation through the builder's own slide API instead.
  def op_builder(op)
    deck = Pptx.build do |d|
      d.slide("Title Slide") { |s| s.title = op["title"] }
      d.slide("Blank") do |s|
        s.shape op["shape"].to_sym, at: [Pptx.inches(1), Pptx.inches(1)], size: [Pptx.inches(3), Pptx.inches(1)],
                                    fill: op["fill"], text: op["text"]
      end
    end
    (@extra_decks ||= []) << deck
  end

  public

  def extra_decks = @extra_decks.to_a
end

# ---------------------------------------------------------------------------
# Validator wrapper
# ---------------------------------------------------------------------------
module Validator
  DOTNET = ENV.fetch("DOTNET", [File.join(Dir.home, ".dotnet", "dotnet"), "/usr/bin/dotnet", "/usr/share/dotnet/dotnet"]
                     .find { |p| File.executable?(p) } || "dotnet")
  DLL = File.join(SKILL_DIR, "bin", "OoxmlValidate.dll")

  module_function

  def ensure_built!
    src = Dir[File.join(SKILL_DIR, "validator", "*.{cs,csproj}")]
    return if File.exist?(DLL) && src.all? { |f| File.mtime(f) <= File.mtime(DLL) }

    warn "building validator (first run downloads DocumentFormat.OpenXml from NuGet)..."
    out, status = Open3.capture2e({ "DOTNET_CLI_TELEMETRY_OPTOUT" => "1", "DOTNET_NOLOGO" => "1" },
                                  DOTNET, "build", "-c", "Release", "-o", File.join(SKILL_DIR, "bin"),
                                  File.join(SKILL_DIR, "validator"))
    abort "validator build failed:\n#{out}" unless status.success?
  end

  def version = ENV.fetch("OOXML_VERSION", "Microsoft365")

  # @return [Array<Hash>] one result per file
  def validate(files)
    return [] if files.empty?
    return files.map { |f| serve(f) } if @server

    ensure_built!
    out, err, status = Open3.capture3(DOTNET, DLL, "--version", version, *files)
    abort "validator failed (#{status.exitstatus}):\n#{err}" if status.exitstatus > 1
    out.lines.map { |l| JSON.parse(l) }
  end

  # Keep one validator process for the block, so many small validations
  # (the minimizer's) do not each pay .NET's startup.
  def session
    ensure_built!
    Open3.popen2(DOTNET, DLL, "--version", version, "--stdin") do |stdin, stdout, _wait|
      @server = [stdin, stdout]
      yield
    ensure
      @server = nil
      stdin.close unless stdin.closed?
    end
  end

  def serve(file)
    stdin, stdout = @server
    stdin.puts(File.expand_path(file))
    stdin.flush
    JSON.parse(stdout.gets || abort("validator process died"))
  end

  # A stable key for "the same bug": error id, element, kind of part, and the
  # description with its concrete values masked.
  def signature(error)
    part = error["part"].to_s.gsub(/\d+(?=\.xml)/, "N").sub(%r{/embeddings/.*}, "/embeddings/*")
    desc = error["description"].gsub(/'[^']*'/, "'…'").gsub(/\b\d+(\.\d+)?\b/, "#")
    # The element and its parent, from the XPath: `radarChart/ser` tells
    # apart bugs that a bare `ser` would lump together.
    where = error["path"].to_s.split("/").last(2).map { |step| step.sub(/\[\d+\]\z/, "").split(":").last }
    where = [error["node"].to_s] if where.empty?
    "#{error['id']} #{part} <#{where.join('/')}> #{desc}"
  end

  def signatures(result)
    return ["OPEN_ERROR #{result['open_error'].gsub(/'[^']*'/, "'…'")}"] if result["open_error"]

    result["errors"].map { |e| signature(e) }.uniq
  end
end

def known_issues
  @known_issues ||= load_known_issues
end

def load_known_issues
  path = File.join(SKILL_DIR, "known-issues.txt")
  return [] unless File.exist?(path)

  File.readlines(path, chomp: true).reject { |l| l.strip.empty? || l.start_with?("#") }.map { |l| Regexp.new(l) }
end

def known?(sig) = known_issues.any? { |re| re.match?(sig) }

def build_case(ops, path)
  ex = Executor.new(ops)
  ex.build
  ex.save(path)
  extras = ex.extra_decks.each_with_index.map do |deck, i|
    extra = path.sub(/\.pptx\z/, "-builder#{i}.pptx")
    deck.save(extra)
    extra
  end
  [ex.log, extras]
end

# ---------------------------------------------------------------------------
# Commands
# ---------------------------------------------------------------------------
def cmd_run(argv)
  opts = { count: 50, seed: Random.new_seed % 2**31, ops: 12, edge: false,
           out: File.join(REPO, "tmp", "ooxml-fuzz") }
  OptionParser.new do |o|
    o.on("--count N", Integer) { opts[:count] = _1 }
    o.on("--seed N", Integer) { opts[:seed] = _1 }
    o.on("--ops N", Integer, "max ops per deck") { opts[:ops] = _1 }
    o.on("--edge", "also generate out-of-range / hostile values") { opts[:edge] = true }
    o.on("--out DIR") { opts[:out] = _1 }
  end.parse!(argv)

  FileUtils.rm_rf(opts[:out])
  FileUtils.mkdir_p(opts[:out])
  master = Random.new(opts[:seed])
  cases = {}
  crashes = []
  rejected = Hash.new(0)

  opts[:count].times do
    seed = master.rand(2**31)
    ops = Generator.new(seed, ops: opts[:ops], edge: opts[:edge]).generate
    path = File.join(opts[:out], "case-#{seed}.pptx")
    begin
      log, extras = build_case(ops, path)
    rescue StandardError => e
      crashes << { "seed" => seed, "stage" => "save", "error" => "#{e.class}: #{e.message[0, 300]}", "backtrace" => e.backtrace.first(8) }
      log = []
      extras = []
    end
    log.select { _1["status"] == "crash" }.each { |c| crashes << c.merge("seed" => seed) }
    log.select { _1["status"] == "rejected" }.each { |r| rejected["#{r['op']}: #{r['error'][0, 90]}"] += 1 }
    File.write(File.join(opts[:out], "case-#{seed}.json"),
               JSON.pretty_generate({ "seed" => seed, "edge" => opts[:edge], "ops" => ops, "log" => log }, allow_nan: true))
    cases[seed] = [path, *extras].select { File.exist?(_1) }
  end

  results = Validator.validate(cases.values.flatten)
  findings = Hash.new { |h, k| h[k] = [] }
  results.each do |r|
    seed = r["file"][/case-(\d+)/, 1].to_i
    Validator.signatures(r).each { |sig| findings[sig] << { "seed" => seed, "file" => r["file"], "example" => r["errors"]&.find { |e| Validator.signature(e) == sig } || r["open_error"] } }
  end

  new_findings = findings.reject { |sig, _| known?(sig) }
  report = { "master_seed" => opts[:seed], "count" => opts[:count], "edge" => opts[:edge],
             "valid" => results.count { _1["ok"] }, "invalid" => results.count { !_1["ok"] },
             "findings" => new_findings.map { |sig, hits| { "signature" => sig, "cases" => hits.size, "first" => hits.first } },
             "known" => findings.select { |sig, _| known?(sig) }.keys,
             "crashes" => crashes, "rejected" => rejected.sort_by { -_2 }.to_h }
  File.write(File.join(opts[:out], "report.json"), JSON.pretty_generate(report, allow_nan: true))

  puts "master seed #{opts[:seed]}#{' (edge)' if opts[:edge]}: #{results.size} decks, " \
       "#{report['valid']} valid, #{report['invalid']} invalid"
  puts "\n#{new_findings.size} distinct validator finding(s):" unless new_findings.empty?
  new_findings.sort_by { -_2.size }.each do |sig, hits|
    ex = hits.first["example"]
    puts "  [#{hits.size}x] #{sig}"
    puts "        e.g. case-#{hits.first['seed']}.json  #{ex['path'] if ex.is_a?(Hash)}"
  end
  puts "\n#{report['known'].size} known issue signature(s) suppressed" unless report["known"].empty?
  unless crashes.empty?
    puts "\n#{crashes.size} crash(es) (non-rejection exceptions):"
    crashes.group_by { _1["error"][0, 120] }.each do |msg, cs|
      lib = cs.first["in_library"] ? "library" : "harness?"
      puts "  [#{cs.size}x, #{lib}] #{msg}\n        e.g. case-#{cs.first['seed']}.json op #{cs.first['i']}  #{cs.first['backtrace']&.first}"
    end
  end
  puts "\nreport: #{File.join(opts[:out], 'report.json')}"
  exit(new_findings.empty? && crashes.none? { _1["in_library"] != false } ? 0 : 1)
end

def load_case(path) = JSON.parse(File.read(path), allow_nan: true)

def cmd_replay(argv)
  keep = nil
  out = nil
  OptionParser.new do |o|
    o.on("--keep LIST", "comma-separated op indexes to keep") { keep = _1.split(",").map(&:to_i) }
    o.on("--out FILE") { out = _1 }
  end.parse!(argv)
  kase = load_case(argv.fetch(0))
  ops = keep ? kase["ops"].values_at(*keep) : kase["ops"]
  out ||= argv[0].sub(/\.json\z/, "-replay.pptx")
  log, extras = build_case(ops, out)
  log.reject { _1["status"] == "ok" && _1["notes"].empty? }.each { puts JSON.generate(_1) }
  Validator.validate([out, *extras]).each { |r| print_result(r) }
end

# Print one validator result, leaving out errors known-issues.txt accepts.
# @return [Boolean] true when nothing but known errors remain
def print_result(result)
  errors = result["errors"].to_a
  known, fresh = errors.partition { |e| known?(Validator.signature(e)) }
  ok = result["open_error"].nil? && fresh.empty?
  note = known.empty? ? "" : " (#{known.size} known error(s) hidden)"
  puts "#{result['file']}: #{ok ? 'valid' : 'INVALID'}#{note}"
  puts "  open error: #{result['open_error']}" if result["open_error"]
  fresh.each { |e| puts "  #{e['id']} #{e['part']} #{e['path']}\n    #{e['description']}" }
  ok
end

# Delta debugging (ddmin) over the op list, keeping any subset whose deck
# still produces +signature+.
def cmd_minimize(argv)
  sig = nil
  OptionParser.new { |o| o.on("--signature SIG", "defaults to the case's first finding") { sig = _1 } }.parse!(argv)
  kase = load_case(argv.fetch(0))
  ops = kase["ops"]
  dir = Dir.mktmpdir("ooxml-min")
  runs = 0

  reproduces = lambda do |subset|
    runs += 1
    path = File.join(dir, "try-#{runs}.pptx")
    _log, extras = build_case(subset, path)
    Validator.validate([path, *extras]).flat_map { Validator.signatures(_1) }.then do |sigs|
      sig ||= sigs.first
      sigs.include?(sig)
    end
  rescue StandardError
    false
  end

  Validator.session do
    abort "case does not reproduce any validator error" unless reproduces.call(ops)
    warn "minimizing for: #{sig}"

    n = 2
    while ops.size >= 2
      chunk = (ops.size / n.to_f).ceil
      subsets = ops.each_slice(chunk).to_a
      reduced = false
      subsets.each_index do |i|
        complement = subsets.reject.with_index { |_, j| j == i }.flatten(1)
        next unless reproduces.call(complement)

        ops = complement
        n = [n - 1, 2].max
        reduced = true
        break
      end
      next if reduced
      break if n >= ops.size

      n = [n * 2, ops.size].min
    end

    # Then shrink inside the surviving ops until nothing more can go: drop any
    # hash key or array element, at any depth, and nil out any value. A
    # removal the op needs makes it fail to build, which simply does not
    # reproduce, so no key has to be marked required.
    loop do
      shrunk = false
      each_shrink(ops) do |trial|
        next unless reproduces.call(trial)

        ops = trial
        shrunk = true
        break
      end
      break unless shrunk
    end

    out = argv[0].sub(/\.json\z/, "-min.json")
    File.write(out, JSON.pretty_generate({ "signature" => sig, "ops" => ops }, allow_nan: true))
    puts "#{runs} builds; #{kase['ops'].size} ops -> #{ops.size}: #{out}"
    puts JSON.pretty_generate(ops, allow_nan: true)
  end
ensure
  FileUtils.rm_rf(dir) if dir
end

# Yield every copy of +value+ with one piece removed or simplified.
def each_shrink(value, &block)
  case value
  when Hash
    value.each_key do |k|
      next if k == "op"

      yield value.except(k)
      yield value.merge(k => nil) unless value[k].nil? || value[k] == false
      each_shrink(value[k]) { |v| yield value.merge(k => v) }
    end
  when Array
    value.each_index do |i|
      yield value[0...i] + value[(i + 1)..]
      each_shrink(value[i]) { |v| yield value.dup.tap { _1[i] = v } }
    end
  end
end

def cmd_validate(argv)
  ok = Validator.validate(argv).map { |r| print_result(r) }
  exit(ok.all? ? 0 : 1)
end

if $PROGRAM_NAME == __FILE__
  cmd = ARGV.shift
  case cmd
  when "run" then cmd_run(ARGV)
  when "replay" then cmd_replay(ARGV)
  when "minimize" then cmd_minimize(ARGV)
  when "validate" then cmd_validate(ARGV)
  else abort File.read(__FILE__).lines.grep(/^#   bundle/).join
  end
end
