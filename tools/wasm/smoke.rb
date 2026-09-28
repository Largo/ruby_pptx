# frozen_string_literal: true

# Run by run.mjs inside ruby.wasm. Fails loudly if Nokogiri is loadable, so
# the run cannot pass by testing the native backend instead.
$stdout.sync = true

begin
  require "nokogiri"
  abort "nokogiri loaded, so this run would not exercise REXML"
rescue LoadError
  nil
end

require "ruby_pptx"
require "stringio"

puts RUBY_DESCRIPTION
abort "expected the REXML backend, got #{Pptx.xml_backend}" unless Pptx.xml_backend == :rexml

prs = Pptx::Presentation.new_default
slide = prs.slides.add(prs.slide_layouts["Title and Content"])
slide.shapes.title.text = "Built in ruby.wasm"
slide.placeholders[1].text_frame.text = "No Nokogiri\nREXML underneath"
slide.notes = "Speaker notes too"
table = slide.shapes.add_table(2, 2, at: [Pptx.inches(1), Pptx.inches(4)],
                                     size: [Pptx.inches(4), Pptx.inches(1)]).table
table[0, 0].text = "Region"
slide.shapes.add_picture("/work/logo.png", at: [Pptx.inches(7), Pptx.inches(0.5)], width: Pptx.inches(1))

data = Pptx::ChartData.new
data.categories = %w[East West]
data.add_series("Q1", [1.5, 2.5])
chart = prs.slides.add(prs.slide_layouts["Blank"]).shapes
           .add_chart(:column_clustered, data, at: [0, 0], size: [Pptx.inches(6), Pptx.inches(4)]).chart
chart.title = "Sales"
prs.save("/work/out.pptx")

back = Pptx::Presentation.open(StringIO.new(File.binread("/work/out.pptx")))
checks = {
  "title" => back.slides[0].shapes.title.text == "Built in ruby.wasm",
  "notes" => back.slides[0].notes == "Speaker notes too",
  "table" => back.slides[0].shapes.find(&:table?).table[0, 0].text == "Region",
  "chart" => back.slides[1].shapes[0].chart.chart_type.name == :COLUMN_CLUSTERED
}
checks.each { |name, ok| puts "#{ok ? "ok" : "FAILED"} #{name}" }
abort "ruby.wasm smoke test failed" unless checks.values.all?
