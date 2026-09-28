# frozen_string_literal: true

# Checks every public member of python-pptx's user-facing API against this gem.
#
#   python3 tools/api_dump.py > api.json
#   ruby -Ilib tools/api_audit.rb api.json
#
# ROADMAP section B once tracked the port module by module, and a module that
# was "done" could still be missing half its members -- speaker notes, picture
# cropping and axis titles all went unnoticed that way. This compares members,
# so "complete" is something that can be checked rather than asserted.
#
# Three tables below decide what counts. Each entry is a claim someone can
# argue with, which is the point of writing them down.

require "json"
require "ruby_pptx"

module ApiAudit
  # python-pptx class => our class. nil means there is no counterpart yet.
  CLASSES = {
    "presentation.Presentation" => "Presentation",
    "slide.Slide" => "Slide", "slide.Slides" => "Slides",
    "slide.SlideLayout" => "SlideLayout", "slide.SlideLayouts" => "SlideLayouts",
    "slide.SlideMaster" => "SlideMaster", "slide.SlideMasters" => "SlideMasters",
    "slide.NotesSlide" => "NotesSlide", "slide.NotesMaster" => "NotesMaster",
    "slide._Background" => "Background",
    "shapes.base.BaseShape" => "BaseShape", "shapes.base._PlaceholderFormat" => "PlaceholderFormat",
    "shapes.autoshape.Shape" => "Shape", "shapes.autoshape.Adjustment" => "Adjustment",
    "shapes.picture.Picture" => "Picture", "shapes.picture.Movie" => "Movie",
    "shapes.connector.Connector" => "Connector", "shapes.graphfrm.GraphicFrame" => "GraphicFrame",
    "shapes.graphfrm._OleFormat" => "OleFormat",
    "shapes.group.GroupShape" => "GroupShape",
    "shapes.freeform.FreeformBuilder" => "FreeformBuilder",
    "shapes.shapetree.SlideShapes" => "SlideShapes", "shapes.shapetree.GroupShapes" => "GroupShapes",
    "shapes.shapetree.LayoutShapes" => "LayoutShapes", "shapes.shapetree.MasterShapes" => "MasterShapes",
    "shapes.shapetree.SlidePlaceholders" => "SlidePlaceholders",
    "shapes.shapetree.LayoutPlaceholders" => "LayoutPlaceholders",
    "shapes.shapetree.MasterPlaceholders" => "MasterPlaceholders",
    "shapes.shapetree.NotesSlideShapes" => "NotesSlideShapes",
    "shapes.shapetree.NotesSlidePlaceholders" => "NotesSlidePlaceholders",
    "shapes.placeholder.SlidePlaceholder" => "SlidePlaceholder",
    "shapes.placeholder.LayoutPlaceholder" => "LayoutPlaceholder",
    "shapes.placeholder.MasterPlaceholder" => "MasterPlaceholder",
    "shapes.placeholder.NotesSlidePlaceholder" => "NotesSlidePlaceholder",
    "shapes.placeholder.PicturePlaceholder" => "PicturePlaceholder",
    "shapes.placeholder.ChartPlaceholder" => "ChartPlaceholder",
    "shapes.placeholder.TablePlaceholder" => "TablePlaceholder",
    "shapes.placeholder.PlaceholderPicture" => "PlaceholderPicture",
    "shapes.placeholder.PlaceholderGraphicFrame" => "PlaceholderGraphicFrame",
    "text.text.TextFrame" => "TextFrame", "text.text._Paragraph" => "Paragraph",
    "text.text._Run" => "Run", "text.text.Font" => "Font",
    "text.text._Hyperlink" => "ActionSetting", "action.Hyperlink" => "ActionSetting",
    "action.ActionSetting" => "ActionSetting",
    "dml.fill.FillFormat" => "FillFormat", "dml.fill._GradientStop" => "GradientStop",
    "dml.line.LineFormat" => "LineFormat", "dml.color.ColorFormat" => "ColorFormat",
    "dml.color.RGBColor" => "RGBColor", "dml.effect.ShadowFormat" => "ShadowFormat",
    "dml.chtfmt.ChartFormat" => "ChartFormat",
    "table.Table" => "Table", "table._Cell" => "Cell", "table._Row" => "TableRow",
    "table._Column" => "TableColumn", "table._Rows" => "TableRows", "table._Columns" => "TableColumns",
    "chart.chart.Chart" => "Chart", "chart.chart.ChartTitle" => "ChartTitle",
    "chart.legend.Legend" => "ChartLegend",
    "chart.axis.CategoryAxis" => "CategoryAxis", "chart.axis.ValueAxis" => "ValueAxis",
    "chart.axis.DateAxis" => "DateAxis", "chart.axis.AxisTitle" => "ChartTitle",
    "chart.axis.TickLabels" => "TickLabels", "chart.axis.MajorGridlines" => "Gridlines",
    "chart.datalabel.DataLabels" => "ChartDataLabels", "chart.datalabel.DataLabel" => "DataLabel",
    "chart.marker.Marker" => "Marker", "chart.point.Point" => "ChartPoint",
    "chart.plot.AreaPlot" => "ChartPlot", "chart.plot.BarPlot" => "ChartPlot",
    "chart.plot.BubblePlot" => "ChartPlot", "chart.plot.DoughnutPlot" => "ChartPlot",
    "chart.plot.LinePlot" => "ChartPlot", "chart.plot.PiePlot" => "ChartPlot",
    "chart.plot.RadarPlot" => "ChartPlot", "chart.plot.XyPlot" => "ChartPlot",
    "chart.series.AreaSeries" => "AreaSeriesView", "chart.series.BarSeries" => "BarSeriesView",
    "chart.series.BubbleSeries" => "BubbleSeriesView", "chart.series.LineSeries" => "LineSeriesView",
    "chart.series.PieSeries" => "PieSeriesView", "chart.series.RadarSeries" => "RadarSeriesView",
    "chart.series.XySeries" => "XySeriesView",
    "chart.category.Categories" => "ChartCategories",
    "chart.data.CategoryChartData" => "ChartData", "chart.data.XyChartData" => "XyChartData",
    "chart.data.BubbleChartData" => "BubbleChartData",
    "chart.data.CategorySeriesData" => "ChartSeries", "chart.data.XySeriesData" => "XySeries",
    "chart.data.BubbleSeriesData" => "BubbleSeries",
    "media.Video" => "Video",
    "chart.plot.Area3DPlot" => "ChartPlot", "chart.data.ChartData" => "ChartData",
    "chart.category.Category" => "ChartCategory",
    "chart.data.Categories" => "ChartDataCategories", "chart.data.Category" => "ChartDataCategory",
    "text.layout.TextFitter" => "TextFitter"
  }.freeze

  # Members that are plumbing rather than API: the lxml element, the owning
  # part, constructors, and things the Ruby redesign has no use for.
  IGNORED = %w[
    element part parent ln get_or_add_ln ph_basename turbo_add_enabled clone_placeholder
    clone_layout_placeholders clone_master_placeholders from_fill_parent from_colorchoice_parent
    from_blob from_path_or_file_like new xml_bytes append data_point_offset series_index
    values_ref x_values_ref y_values_ref bubble_sizes_ref name_ref number_format categories_ref
    chart_part iter_values
    ph_type orient sz idx
  ].freeze

  # python-pptx member => the names the Ruby API uses for it.
  RENAMES = {
    "add_slide" => %w[add], "get" => %w[by_id by_idx by_type []], "get_by_name" => %w[by_name []],
    "remove" => %w[delete], "slide_layout" => %w[layout], "chart_title" => %w[title],
    "horz_banding" => %w[banded_rows], "vert_banding" => %w[banded_columns],
    "language_id" => %w[language], "iter_cells" => %w[each_cell],
    "iter_cloneable_placeholders" => %w[cloneable_placeholders],
    "follow_master_background" => %w[follows_master_background?],
    "is_placeholder" => %w[placeholder?], "is_merge_origin" => %w[merge_origin?],
    "is_spanned" => %w[spanned?], "has_text_frame" => %w[text_frame?], "has_chart" => %w[chart?],
    "has_table" => %w[table?], "has_notes_slide" => %w[notes_slide?],
    "add_category" => %w[categories=],
    "slide_master" => %w[slide_master slide_masters],
    "number_format_is_linked" => %w[number_format_linked?], "axis_title" => %w[title],
    "shape_offset_x" => %w[offset_x], "shape_offset_y" => %w[offset_y],
    "add_data_point" => %w[<< add_point add_series],
    # python-pptx wraps the address in a Hyperlink object; here it is flattened.
    "hyperlink" => %w[address url hyperlink]
  }.freeze

  # Classes that are public by name upstream but that its API never hands
  # out, so there is nothing for a caller to reach.
  INTERNAL_CLASSES = {
    # AdjustmentCollection#__getitem__ returns plain floats, never these.
    "shapes.autoshape.Adjustment" => "internal to AdjustmentCollection",
    # Returned by add_data_point, but nothing reads them back.
    "chart.data.CategoryDataPoint" => "value holder", "chart.data.XyDataPoint" => "value holder",
    "chart.data.BubbleDataPoint" => "value holder",
    # The lookup table behind add_shape; here Pptx::AutoShapeSpec.
    "shapes.autoshape.AutoShapeType" => "lookup table",
    # Deprecated upstream in favour of placeholder_format.
    "shapes.placeholder.BasePlaceholder" => "deprecated",
    # Abstract bases and typing protocols.
    "shapes.shapetree.BasePlaceholders" => "abstract", "shared.ElementProxy" => "abstract",
    "shared.ParentedElementProxy" => "abstract", "shared.PartElementProxy" => "abstract",
    "types.ProvidesExtents" => "typing protocol", "types.ProvidesPart" => "typing protocol",
    # Font discovery by family name. It supports only macOS and Windows upstream
    # and raises on Linux; this gem takes font_file: instead. See PORTING.md.
    "text.fonts.FontFiles" => "not ported by design"
  }.freeze

  # Whole families that are not API: private helpers (a leading underscore on
  # the class name), writers and rewriters, and the data-point value
  # objects python-pptx builds internally from what add_series is given.
  UNCHECKED = [
    /\._[A-Z]/, /\.chart\.xmlwriter\./, /\.chart\.xlsx\./, /WorkbookWriter\z/,
    /\Apackage\./, /\Ashapes\.Subshape\z/, /\Achart\.plot\.PlotTypeInspector\z/,
    /\Amedia\.Video\z/
  ].freeze

  # Members that exist on a python-pptx class only by inheritance and mean
  # nothing for it -- a category series has no x values.
  IGNORED_PER_CLASS = {
    "chart.data.CategorySeriesData" => %w[x_values y_values]
  }.freeze

  # Classes where only some members are API. A category label upstream is a
  # str subclass, so it also "has" every string method.
  ONLY_MEMBERS = {
    "chart.category.Category" => %w[idx label]
  }.freeze

  module_function

  def candidates(member)
    stem = member.sub(/\A(has|is)_/, "")
    [member, "#{member}?", "#{stem}?", stem, *RENAMES.fetch(member, [])].uniq
  end

  def ours(name)
    Pptx.const_get(name) if name && Pptx.const_defined?(name, false)
  end

  # @return [Hash] {missing_classes: [...], missing_members: {pyclass => [..]}}
  def run(dump)
    classes = dump.fetch("classes")
    checked = CLASSES.reject { |pyclass, _| INTERNAL_CLASSES.key?(pyclass) || !classes.key?(pyclass) }
    missing_classes = checked.filter_map do |pyclass, rbname|
      "#{pyclass} (-> Pptx::#{rbname})" if ours(rbname).nil?
    end
    missing_members = checked.each_with_object({}) do |(pyclass, rbname), result|
      klass = ours(rbname) or next
      gaps = member_gaps(pyclass, classes[pyclass], klass)
      result[pyclass] = gaps unless gaps.empty?
    end
    { "version" => dump["version"], "missing_classes" => missing_classes,
      "missing_members" => missing_members, "unaccounted" => unaccounted(classes.keys) }
  end

  # Members of +pyclass+ with no counterpart on +klass+ under any known name.
  def member_gaps(pyclass, members, klass)
    # Class methods count too: `RGBColor.from_string` is one on both sides.
    have = (klass.public_instance_methods + klass.singleton_methods).map(&:to_s)
    members &= ONLY_MEMBERS[pyclass] if ONLY_MEMBERS.key?(pyclass)
    (members - IGNORED - IGNORED_PER_CLASS.fetch(pyclass, [])).reject do |member|
      candidates(member).any? { |c| have.include?(c) || have.include?("#{c}=") }
    end
  end

  # Upstream classes neither checked nor excused.
  def unaccounted(names)
    names.reject do |name|
      CLASSES.key?(name) || INTERNAL_CLASSES.key?(name) || UNCHECKED.any? { |pattern| pattern.match?(name) }
    end.sort
  end
end

if $PROGRAM_NAME == __FILE__
  result = ApiAudit.run(JSON.parse(File.read(ARGV.fetch(0))))
  if ARGV.include?("--json")
    puts JSON.pretty_generate(result)
  else
    puts "python-pptx #{result["version"]}: members with no counterpart in ruby_pptx"
    puts
    result["missing_classes"].each { |c| puts "  class  #{c}" }
    result["missing_members"].sort.each { |c, ms| puts format("  %-36s %s", c, ms.join(" ")) }
    unless result["unaccounted"].empty?
      puts "\nclasses the audit neither checks nor excuses -- map them or say why not:"
      result["unaccounted"].each { |c| puts "  #{c}" }
    end
    total = result["missing_classes"].size + result["missing_members"].values.sum(&:size)
    puts "\n#{total.zero? ? "nothing missing" : "#{total} gaps"}"
  end
end
