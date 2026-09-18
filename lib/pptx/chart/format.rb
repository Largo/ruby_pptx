# frozen_string_literal: true

require "pptx/element_proxy"
require "pptx/text/text"
require "pptx/enum/chart"
require "pptx/dml/fill"

module Pptx
  # The fill and outline of a chart element -- a series, an axis, gridlines.
  #
  # Reached through the `format` of whatever it belongs to.
  class ChartFormat < ElementProxy
    def fill = @fill ||= FillFormat.from_fill_parent(@element.get_or_add_spPr)

    def line = @line ||= LineFormat.new(@element.get_or_add_spPr)

    def inspect = "#<Pptx::ChartFormat #{fill.type&.name}>"
  end

  # A chart's legend.
  class ChartLegend < ElementProxy
    # @return [Pptx::Enum::XL_LEGEND_POSITION, nil] nil when PowerPoint decides
    def position
      value = @element.legendPos&.val
      value && Enum::XL_LEGEND_POSITION.from_xml(value)
    end

    def position=(value)
      @element.get_or_add_legendPos.val =
        Enum::XL_LEGEND_POSITION.to_xml(Enum::XL_LEGEND_POSITION.fetch(value))
      value
    end

    # True when the legend sits inside the plot area rather than beside it.
    def include_in_layout? = @element.overlay&.val || false

    def include_in_layout=(value)
      @element.get_or_add_overlay.val = value
      value
    end

    def inspect = "#<Pptx::ChartLegend #{position&.name}>"
  end

  # A chart or axis title.
  #
  # The title carries a text frame like any shape, so it can be formatted the
  # same way.
  class ChartTitle < ElementProxy
    def text_frame = TextFrame.new(@element.rich, self)

    # Titles are formatting, not content; there is no part to reach for.
    def part = nil

    def text = text_frame.text

    def text=(value)
      text_frame.text = value
      value
    end

    def inspect = "#<Pptx::ChartTitle #{text.inspect}>"
  end

  # One axis of a chart.
  class ChartAxis < ElementProxy
    # Whether the axis is drawn. A `c:delete` of 1 hides it.
    def visible? = !(@element.delete&.val || false)

    def visible=(value)
      @element.get_or_add_delete.val = !value
      value
    end

    def major_gridlines? = !@element.majorGridlines.nil?

    def major_gridlines=(value)
      value ? @element.get_or_add_majorGridlines : @element.remove_majorGridlines
      value
    end

    def minor_gridlines? = !@element.minorGridlines.nil?

    def minor_gridlines=(value)
      value ? @element.get_or_add_minorGridlines : @element.remove_minorGridlines
      value
    end

    # The fixed end of the scale, or nil when PowerPoint scales automatically.
    def maximum_scale = @element.scaling.maximum

    def maximum_scale=(value)
      @element.scaling.maximum = value
      value
    end

    def minimum_scale = @element.scaling.minimum

    def minimum_scale=(value)
      @element.scaling.minimum = value
      value
    end

    def major_unit = @element.majorUnit&.val

    def major_unit=(value)
      value.nil? ? @element.remove_majorUnit : (@element.get_or_add_majorUnit.val = value)
      value
    end

    def minor_unit = @element.minorUnit&.val

    def minor_unit=(value)
      value.nil? ? @element.remove_minorUnit : (@element.get_or_add_minorUnit.val = value)
      value
    end

    # The number format of the tick labels, e.g. "0.0%".
    def number_format = @element.numFmt&.formatCode

    def number_format=(value)
      format = @element.get_or_add_numFmt
      format.formatCode = value
      # An explicit format is no longer taken from the source data.
      format.sourceLinked = false
      value
    end

    def has_title? = !@element.title.nil?

    def title
      @element.title.nil? ? nil : ChartTitle.new(@element.title)
    end

    # Give the axis a title, or remove it with nil.
    def title=(text)
      if text.nil?
        @element.remove_title
        return nil
      end

      element = @element.title || begin
        created = Oxml::CT_Title.new_title(@element)
        @element.insert_title(created)
        created
      end
      ChartTitle.new(element).text = text
    end

    def inspect = "#<Pptx::ChartAxis #{@element.nsptag} visible=#{visible?}>"
  end

  # One plot -- a "chart group" in the MS API -- within a chart.
  #
  # A chart usually has exactly one; a combo chart has several, which is why
  # these are a collection rather than properties of the chart itself.
  class ChartPlot < ElementProxy
    def initialize(element, chart)
      super(element)
      @chart = chart
    end

    # The space between bars, as a percentage of bar width.
    def gap_width = @element.gapWidth&.val || 150

    def gap_width=(value)
      @element.get_or_add_gapWidth.val = value
      value
    end

    # How far bars in a group overlap, -100 to 100.
    def overlap = @element.overlap&.val || 0

    def overlap=(value)
      @element.get_or_add_overlap.val = value
      value
    end

    # Whether each data point gets its own colour, as a pie does.
    def vary_by_categories? = @element.varyColors&.val || false

    def vary_by_categories=(value)
      @element.get_or_add_varyColors.val = value
      value
    end

    # Whether this plot shows data labels at all.
    def data_labels? = !@element.dLbls.nil?

    def data_labels=(value)
      value ? @element.get_or_add_default_dLbls : @element.remove_dLbls
      @data_labels = nil
      value
    end

    # The data-label settings, switched on with PowerPoint's defaults if the
    # plot has none yet. python-pptx makes you set `has_data_labels` first;
    # asking for them is enough here.
    def data_labels = @data_labels ||= ChartDataLabels.new(@element.get_or_add_default_dLbls)

    def inspect = "#<Pptx::ChartPlot #{@element.nsptag}>"
  end

  # The data labels of a plot.
  class ChartDataLabels < ElementProxy
    SHOW_FLAGS = {
      value: "c:showVal", category_name: "c:showCatName", series_name: "c:showSerName",
      percentage: "c:showPercent", legend_key: "c:showLegendKey",
      bubble_size: "c:showBubbleSize"
    }.freeze

    SHOW_FLAGS.each do |name, tag|
      local = Oxml::Ns.split_tag(tag).last

      define_method("show_#{name}?") { @element.public_send(local)&.val || false }

      define_method("show_#{name}=") do |value|
        @element.public_send("get_or_add_#{local}").val = value
        value
      end
    end

    def number_format = @element.numFmt&.formatCode

    def number_format=(value)
      format = @element.get_or_add_numFmt
      format.formatCode = value
      format.sourceLinked = false
      value
    end

    # @return [Pptx::Enum::XL_LABEL_POSITION, nil]
    def position
      value = @element.dLblPos&.val
      value && Enum::XL_LABEL_POSITION.from_xml(value)
    end

    def position=(value)
      if value.nil?
        @element.remove_dLblPos
        return nil
      end

      @element.get_or_add_dLblPos.val =
        Enum::XL_LABEL_POSITION.to_xml(Enum::XL_LABEL_POSITION.fetch(value))
      value
    end

    def inspect = "#<Pptx::ChartDataLabels value=#{show_value?}>"
  end
end
