# frozen_string_literal: true

require "ruby_pptx/element_proxy"
require "ruby_pptx/sliceable"
require "ruby_pptx/text/text"
require "ruby_pptx/enum/chart"
require "ruby_pptx/dml/fill"

module Pptx
  # The fill and outline of a chart element -- a series, a point, an axis,
  # gridlines, a title or a marker.
  #
  # Reached through the `format` of whatever it belongs to.
  class ChartFormat < ElementProxy
    def fill
      @fill ||= FillFormat.from_fill_parent(@element.get_or_add_spPr)
    end

    def line
      @line ||= LineFormat.new(@element.get_or_add_spPr)
    end

    def inspect
      "#<Pptx::ChartFormat #{fill.type&.name}>"
    end
  end

  # A chart's legend.
  class ChartLegend < ElementProxy
    # The legend's font, created on first use.
    def font
      @font ||= Font.new(@element.defRPr)
    end

    # Where the legend sits. Absent from the file means PowerPoint's default,
    # which is the right-hand side.
    #
    # @return [Pptx::Enum::Member] a member of XL_LEGEND_POSITION
    def position
      @element.legendPos&.val || Enum::XL_LEGEND_POSITION::RIGHT
    end

    def position=(value)
      @element.get_or_add_legendPos.val = Enum::XL_LEGEND_POSITION.fetch(value)
    end

    # How far the legend is shifted sideways, as a fraction of the chart
    # width: -1.0 to 1.0, with 0.0 meaning PowerPoint places it.
    def horz_offset
      @element.horz_offset
    end

    def horz_offset=(value)
      @element.horz_offset = value
    end

    # True when the legend sits inside the plot area, overlapping it, rather
    # than having space made for it. An absent `c:overlay` reads as true.
    def include_in_layout?
      @element.overlay.nil? || @element.overlay.val
    end

    # nil removes the setting and returns to the default.
    def include_in_layout=(value)
      if value.nil?
        @element.remove_overlay
      else
        @element.get_or_add_overlay.val = value ? true : false
      end
    end

    def inspect
      "#<Pptx::ChartLegend #{position.name}>"
    end
  end

  # A chart or axis title.
  #
  # The title carries a text frame like any shape, so it can be formatted the
  # same way.
  class ChartTitle < ElementProxy
    # The title box's fill and outline.
    def format
      @format ||= ChartFormat.new(@element)
    end

    # True when the title has text of its own rather than being generated.
    def text_frame?
      !@element.rich.nil?
    end

    # The title's text, created on first use.
    def text_frame
      TextFrame.new(@element.get_or_add_rich, self)
    end

    # Titles are formatting, not content; there is no part to reach for.
    def part
      nil
    end

    def text
      text_frame? ? text_frame.text : ""
    end

    def text=(value)
      text_frame.text = value
    end

    def inspect
      "#<Pptx::ChartTitle #{text.inspect}>"
    end
  end

  # The major gridlines of an axis.
  class Gridlines < ElementProxy
    def format
      @format ||= ChartFormat.new(@element.get_or_add_majorGridlines)
    end
  end

  # The labels along an axis: their font, number format and spacing.
  class TickLabels < ElementProxy
    def font
      @font ||= Font.new(@element.defRPr)
    end

    # The Excel number format, "General" when none is set.
    def number_format
      @element.numFmt&.formatCode || "General"
    end

    # Setting a format stops it following the source data's format.
    def number_format=(value)
      @element.get_or_add_numFmt.formatCode = value
      self.number_format_linked = false
    end

    # True when the labels take their number format from the worksheet.
    def number_format_linked?
      format = @element.numFmt
      return false if format.nil?

      format.sourceLinked.nil? || format.sourceLinked
    end

    def number_format_linked=(value)
      @element.get_or_add_numFmt.sourceLinked = value
    end

    # Distance of the labels from the axis as a percentage of the default;
    # 100 is the default. Only a category axis has one.
    def offset
      @element.respond_to?(:lblOffset) ? (@element.lblOffset&.val || 100) : 100
    end

    def offset=(value)
      raise Error, "only a category axis has a label offset" unless @element.nsptag == "c:catAx"

      @element.remove_lblOffset
      return if value == 100

      @element.get_or_add_lblOffset.val = value
    end
  end

  # One axis of a chart. {CategoryAxis}, {DateAxis} and {ValueAxis} add what
  # only they have.
  class ChartAxis < ElementProxy
    # The axis line and its fill.
    def format
      @format ||= ChartFormat.new(@element)
    end

    # Whether the axis is drawn. A `c:delete` of true hides it.
    #
    # An axis with no `c:delete` at all is drawn, which is what the schema and
    # PowerPoint say. python-pptx reports it as hidden; see PORTING.md.
    def visible?
      @element.delete.nil? || !@element.delete.val
    end

    def visible=(value)
      unless [true, false].include?(value)
        raise ArgumentError, "visible must be true or false, got #{value.inspect}"
      end

      @element.get_or_add_delete.val = !value
    end

    def major_gridlines?
      !@element.majorGridlines.nil?
    end

    def major_gridlines=(value)
      value ? @element.get_or_add_majorGridlines : @element.remove_majorGridlines
    end

    # The major gridlines, for formatting. Asking for them switches them on.
    def major_gridlines
      @major_gridlines ||= Gridlines.new(@element)
    end

    def minor_gridlines?
      !@element.minorGridlines.nil?
    end

    def minor_gridlines=(value)
      value ? @element.get_or_add_minorGridlines : @element.remove_minorGridlines
    end

    # @return [Pptx::Enum::Member] a member of XL_TICK_MARK
    def major_tick_mark
      @element.majorTickMark&.val || Enum::XL_TICK_MARK::CROSS
    end

    # The default, cross, is written by leaving the element out.
    def major_tick_mark=(value)
      set_tick_mark(:majorTickMark, value)
    end

    def minor_tick_mark
      @element.minorTickMark&.val || Enum::XL_TICK_MARK::CROSS
    end

    def minor_tick_mark=(value)
      set_tick_mark(:minorTickMark, value)
    end

    # The fixed end of the scale, or nil when PowerPoint scales automatically.
    def maximum_scale
      @element.scaling.maximum
    end

    def maximum_scale=(value)
      @element.scaling.maximum = value
    end

    def minimum_scale
      @element.scaling.minimum
    end

    def minimum_scale=(value)
      @element.scaling.minimum = value
    end

    # True when the axis runs from its maximum to its minimum.
    def reverse_order?
      @element.orientation == Oxml::SimpleTypes::ST_Orientation::MAX_MIN
    end

    def reverse_order=(value)
      @element.orientation = if value
                               Oxml::SimpleTypes::ST_Orientation::MAX_MIN
                             else
                               Oxml::SimpleTypes::ST_Orientation::MIN_MAX
                             end
    end

    # @return [Pptx::Enum::Member] a member of XL_TICK_LABEL_POSITION
    def tick_label_position
      @element.tickLblPos&.val || Enum::XL_TICK_LABEL_POSITION::NEXT_TO_AXIS
    end

    def tick_label_position=(value)
      @element.get_or_add_tickLblPos.val = Enum::XL_TICK_LABEL_POSITION.fetch(value)
    end

    # The labels along the axis.
    def tick_labels
      @tick_labels ||= TickLabels.new(@element)
    end

    # Shortcuts for the tick-label number format, the setting most often
    # wanted from an axis.
    def number_format
      tick_labels.number_format
    end

    def number_format=(value)
      tick_labels.number_format = value
    end

    def title?
      !@element.title.nil?
    end

    # @return [ChartTitle, nil] nil when the axis has no title
    def title
      element = @element.title
      element && ChartTitle.new(element)
    end

    # A String sets the title's text; true shows a title with no text of its
    # own; nil or false removes it.
    def title=(value)
      case value
      when nil, false then @element.remove_title
      when true then @element.get_or_add_title
      else ChartTitle.new(@element.get_or_add_title).text = value
      end
    end

    def inspect
      "#<#{self.class.name} visible=#{visible?}>"
    end

    private

    def set_tick_mark(name, value)
      member = Enum::XL_TICK_MARK.fetch(value)
      @element.public_send(:"remove_#{name}")
      return if member == Enum::XL_TICK_MARK::CROSS

      @element.public_send(:"get_or_add_#{name}").val = member
    end
  end

  # The horizontal axis of most charts: one label per category.
  class CategoryAxis < ChartAxis
    def category_type
      Enum::XL_CATEGORY_TYPE::CATEGORY_SCALE
    end
  end

  # A category axis whose categories are dates.
  class DateAxis < ChartAxis
    def category_type
      Enum::XL_CATEGORY_TYPE::TIME_SCALE
    end
  end

  # An axis measuring values. An XY or bubble chart has two.
  class ValueAxis < ChartAxis
    # Where the *other* axis crosses this one. `crosses` describes the
    # crossing and is stored on the axis being crossed, which is why these
    # two read and write the perpendicular axis.
    #
    # @return [Pptx::Enum::Member] a member of XL_AXIS_CROSSES; CUSTOM when a
    #   specific value is set through {#crosses_at}
    def crosses
      cross_axis.crosses&.val || Enum::XL_AXIS_CROSSES::CUSTOM
    end

    def crosses=(value)
      member = Enum::XL_AXIS_CROSSES.fetch(value)
      axis = cross_axis
      return if member == Enum::XL_AXIS_CROSSES::CUSTOM && axis.crossesAt

      axis.remove_crosses
      axis.remove_crossesAt
      if member == Enum::XL_AXIS_CROSSES::CUSTOM
        axis.get_or_add_crossesAt.val = 0.0
      else
        axis.get_or_add_crosses.val = member
      end
    end

    # The value on this axis at which the other crosses, or nil.
    def crosses_at
      cross_axis.crossesAt&.val
    end

    def crosses_at=(value)
      axis = cross_axis
      axis.remove_crosses
      axis.remove_crossesAt
      axis.get_or_add_crossesAt.val = value unless value.nil?
    end

    # The distance between major tick marks, or nil when automatic.
    def major_unit
      @element.majorUnit&.val
    end

    def major_unit=(value)
      @element.remove_majorUnit
      @element.get_or_add_majorUnit.val = value unless value.nil?
    end

    def minor_unit
      @element.minorUnit&.val
    end

    def minor_unit=(value)
      @element.remove_minorUnit
      @element.get_or_add_minorUnit.val = value unless value.nil?
    end

    private

    # The axis whose `c:axId` this one names in its `c:crossAx`.
    def cross_axis
      id = @element.crossAx.val
      # A union of whole paths rather than `(a | b)/c`: the same query, but
      # REXML's XPath does not accept a union as a path step.
      query = %w[c:catAx c:valAx c:dateAx].map { |axis| "../#{axis}/c:axId[@val=\"#{id}\"]" }.join(" | ")
      @element.xpath(query).first.parent
    end
  end

  # The data labels of a plot or series.
  class ChartDataLabels < ElementProxy
    SHOW_FLAGS = {
      value: "c:showVal", category_name: "c:showCatName", series_name: "c:showSerName",
      percentage: "c:showPercent", legend_key: "c:showLegendKey",
      bubble_size: "c:showBubbleSize"
    }.freeze

    # Reading a flag does not add it. python-pptx's getters do -- asking
    # whether values are shown writes `c:showVal val="0"` into the file --
    # and a read that changes the document is not one this library will make.
    SHOW_FLAGS.each do |name, tag|
      local = Oxml::Ns.split_tag(tag).last

      define_method("show_#{name}?") do
        flag = @element.public_send(local)
        flag.nil? ? false : flag.val
      end

      define_method("show_#{name}=") do |value|
        @element.public_send("get_or_add_#{local}").val = value ? true : false
      end
    end

    def font
      @font ||= Font.new(@element.defRPr)
    end

    # The Excel number format, "General" when none is set.
    def number_format
      @element.numFmt&.formatCode || "General"
    end

    def number_format=(value)
      @element.get_or_add_numFmt.formatCode = value
      self.number_format_linked = false
    end

    # True when the labels take their number format from the worksheet. Data
    # labels with no number format at all follow the source.
    def number_format_linked?
      format = @element.numFmt
      format.nil? || format.sourceLinked.nil? || format.sourceLinked
    end

    def number_format_linked=(value)
      @element.get_or_add_numFmt.sourceLinked = value
    end

    # @return [Pptx::Enum::Member, nil] a member of XL_DATA_LABEL_POSITION
    def position
      @element.dLblPos&.val
    end

    def position=(value)
      if value.nil?
        @element.remove_dLblPos
      else
        @element.get_or_add_dLblPos.val = Enum::XL_DATA_LABEL_POSITION.fetch(value)
      end
    end

    def inspect
      "#<Pptx::ChartDataLabels value=#{show_value?}>"
    end
  end

  # The label of a single data point, overriding the series' labels for it.
  class DataLabel
    def initialize(ser, idx)
      @ser = ser
      @idx = idx
    end

    def font
      @font ||= TextFrame.new(label_element.get_or_add_txPr, self).paragraphs[0].font
    end

    # True when the label has text of its own rather than a generated value.
    def text_frame?
      !@ser.dLbl_for_point(@idx)&.rich.nil?
    end

    # The label's own text, replacing the generated value. Created on use.
    def text_frame
      TextFrame.new(label_element.get_or_add_rich, self)
    end

    # Labels are formatting, not content; there is no part to reach for.
    def part
      nil
    end

    # @return [Pptx::Enum::Member, nil] a member of XL_DATA_LABEL_POSITION
    def position
      @ser.dLbl_for_point(@idx)&.dLblPos&.val
    end

    def position=(value)
      if value.nil?
        @ser.dLbl_for_point(@idx)&.remove_dLblPos
      else
        label_element.get_or_add_dLblPos.val = Enum::XL_DATA_LABEL_POSITION.fetch(value)
      end
    end

    def inspect
      "#<Pptx::DataLabel point=#{@idx}>"
    end

    private

    # This point's `c:dLbl`, created on first use.
    def label_element
      @ser.get_or_add_dLbl_for_point(@idx)
    end
  end

  # The symbol drawn at each point of a line, radar or XY series, or at one
  # point.
  class Marker < ElementProxy
    def format
      @format ||= ChartFormat.new(@element.get_or_add_marker)
    end

    # Size in points, 2 to 72, or nil when inherited.
    def size
      @element.marker&.size&.val
    end

    def size=(value)
      marker = @element.get_or_add_marker
      marker.remove_size
      marker.get_or_add_size.val = value unless value.nil?
    end

    # @return [Pptx::Enum::Member, nil] a member of XL_MARKER_STYLE
    def style
      @element.marker&.symbol&.val
    end

    def style=(value)
      marker = @element.get_or_add_marker
      marker.remove_symbol
      marker.get_or_add_symbol.val = Enum::XL_MARKER_STYLE.fetch(value) unless value.nil?
    end
  end

  # One data point of a series, for formatting it apart from the rest.
  class ChartPoint
    def initialize(ser, idx)
      @ser = ser
      @idx = idx
    end

    def data_label
      @data_label ||= DataLabel.new(@ser, @idx)
    end

    def format
      @format ||= ChartFormat.new(@ser.get_or_add_dPt_for_point(@idx))
    end

    def marker
      @marker ||= Marker.new(@ser.get_or_add_dPt_for_point(@idx))
    end

    def inspect
      "#<Pptx::ChartPoint #{@idx}>"
    end
  end

  # The points of a series, indexed from 0.
  class ChartPoints
    include Enumerable
    include Sliceable

    def initialize(ser, count)
      @ser = ser
      @count = count
    end

    def size
      @count
    end
    alias length size

    def [](index, length = nil)
      slice_members((0...@count).to_a, index, length) { |i| ChartPoint.new(@ser, i) }
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @count.times { |i| yield ChartPoint.new(@ser, i) }
      self
    end

    def inspect
      "#<Pptx::ChartPoints size=#{size}>"
    end
  end

  # One plot -- a "chart group" in the MS API -- within a chart.
  #
  # A chart usually has exactly one; a combo chart has several, which is why
  # these are a collection rather than properties of the chart itself.
  class ChartPlot < ElementProxy
    attr_reader :chart

    def initialize(element, chart)
      super(element)
      @chart = chart
    end

    # The categories this plot's series are plotted against.
    def categories
      ChartCategories.new(@element)
    end

    # This plot's series, in the order the chart draws them.
    def series
      @element.sers.map { |ser| ChartSeriesView.for(ser, @chart) }
    end

    # The space between bars, as a percentage of bar width.
    def gap_width
      @element.gapWidth&.val || 150
    end

    def gap_width=(value)
      @element.get_or_add_gapWidth.val = value
    end

    # How far bars in a group overlap, -100 to 100. Zero is written by
    # leaving the element out.
    def overlap
      @element.overlap&.val || 0
    end

    def overlap=(value)
      if value.zero?
        @element.remove_overlap
      else
        @element.get_or_add_overlap.val = value
      end
    end

    # The bubble size as a percentage of the default, 0 to 300.
    def bubble_scale
      @element.bubbleScale&.val || 100
    end

    # nil returns to the default.
    def bubble_scale=(value)
      @element.remove_bubbleScale
      @element.get_or_add_bubbleScale.val = value unless value.nil?
    end

    # Whether each data point gets its own colour, as a pie does. An absent
    # `c:varyColors` reads as true, the schema default.
    def vary_by_categories?
      @element.varyColors.nil? || @element.varyColors.val
    end

    def vary_by_categories=(value)
      @element.get_or_add_varyColors.val = value ? true : false
    end

    # Whether this plot shows data labels at all.
    def data_labels?
      !@element.dLbls.nil?
    end

    # Switching labels on shows values, as PowerPoint's default does.
    def data_labels=(value)
      if value
        @element.get_or_add_default_dLbls.showVal.val = true unless @element.dLbls
      else
        @element.remove_dLbls
      end
      @data_labels = nil
    end

    # The data-label settings, switched on with PowerPoint's defaults if the
    # plot has none yet. python-pptx makes you set `has_data_labels` first;
    # asking for them is enough here.
    def data_labels
      @data_labels ||= ChartDataLabels.new(@element.get_or_add_default_dLbls)
    end

    def inspect
      "#<Pptx::ChartPlot #{@element.nsptag}>"
    end
  end

  # The categories of a plot, read back from the chart's cached values.
  #
  # A category can be empty -- an empty worksheet cell -- in which case it
  # has no cached point but still counts, so {#size} is the declared count
  # rather than the number of labels found.
  class ChartCategories
    include Enumerable

    def initialize(plot_element)
      @plot = plot_element
    end

    def each
      return enum_for(:each) { size } unless block_given?

      @plot.cat_pts.each_with_index { |point, i| yield ChartCategory.new(point, i) }
      self
    end

    def [](index)
      to_a[index]
    end

    def size
      @plot.cat_pt_count
    end
    alias length size

    # How many levels of labels there are: 0 with no categories, 1 for a
    # plain list, more for grouped categories.
    def depth
      cat = @plot.cat
      return 0 if cat.nil?
      return 1 if cat.multiLvlStrRef.nil?

      cat.lvls.size
    end

    # Each level's labels, leaf level first.
    def levels
      cat = @plot.cat
      return [] if cat.nil?

      cat.lvls.map { |lvl| lvl.xpath("./c:pt").map { |pt| ChartCategory.new(pt) } }
    end

    # One tuple per leaf category, its labels listed root first -- the form
    # ChartData takes them in.
    def flattened_labels
      return [] if @plot.cat.nil?
      return map { |category| [category.label] } if @plot.cat.multiLvlStrRef.nil?

      leaf, *parents = levels
      leaf.map { |category| lineage(category, parents).reverse.map(&:label) }
    end

    def inspect
      "#<Pptx::ChartCategories #{map(&:label).inspect}>"
    end

    private

    # The category followed by its parent at each level above it: the last
    # parent whose index is not past the child's.
    def lineage(category, parents)
      parents.each_with_object([category]) do |level, chain|
        break chain if level.empty?

        chain << (level.take_while { |parent| parent.idx <= chain.last.idx }.last || level.first)
      end
    end
  end

  # One category label.
  class ChartCategory
    def initialize(point, idx = nil)
      @point = point
      @idx = idx
    end

    # The label, "" for an empty category.
    def label
      @point.nil? ? "" : @point.v.text
    end

    def idx
      @point.nil? ? @idx : @point.idx
    end

    def to_s
      label
    end
    alias to_str to_s

    def ==(other)
      other.is_a?(ChartCategory) ? label == other.label && idx == other.idx : label == other
    end

    def inspect
      "#<Pptx::ChartCategory #{idx}:#{label.inspect}>"
    end
  end
end
