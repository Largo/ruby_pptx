# frozen_string_literal: true

require "ruby_pptx/chart/data"

module Pptx
  # Data for a scatter chart: series of (x, y) points.
  #
  #   data = Pptx::XyChartData.new
  #   data.add_series("Alpha") do |s|
  #     s.add_point(1, 10)
  #     s.add_point(2, 20)
  #   end
  #
  # Unlike a category chart, each series has its own x values, so the
  # worksheet stacks the series in blocks rather than putting them side by
  # side in shared columns.
  class XyChartData
    include Enumerable

    WORKSHEET_NAME = "Sheet1"
    DEFAULT_NUMBER_FORMAT = "General"
    # Each series block is followed by a blank row, and preceded by its title
    # row -- two rows of overhead per series.
    ROWS_PER_SERIES_OVERHEAD = 2

    attr_reader :series, :number_format

    def initialize(number_format: DEFAULT_NUMBER_FORMAT)
      @series = []
      @number_format = number_format
    end

    # Add a series. Points may be given up front, added in a block, or both.
    #
    # @param points [Array<Array>] each [x, y]
    # @return [XySeries]
    def add_series(name, points: [], number_format: nil)
      series = series_class.new(self, @series.size, name, number_format || @number_format)
      @series << series
      points.each { |point| series.add_point(*point) }
      yield series if block_given?
      series
    end

    def each(&)
      @series.each(&)
    end

    def size
      @series.size
    end

    # The row a series' title sits on, counting the blocks before it.
    #
    # @api private
    def title_row(series)
      preceding_points = @series[0...series.index].sum(&:size)
      (series.index * ROWS_PER_SERIES_OVERHEAD) + preceding_points + 1
    end

    # @api private
    def series_name_ref(series)
      "#{WORKSHEET_NAME}!$B$#{title_row(series)}"
    end

    # @api private
    def x_values_ref(series)
      column_ref(series, "A")
    end

    # @api private
    def y_values_ref(series)
      column_ref(series, "B")
    end

    # @api private
    def column_ref(series, column)
      first = title_row(series) + 1
      last = first + [series.size, 1].max - 1
      "#{WORKSHEET_NAME}!$#{column}$#{first}:$#{column}$#{last}"
    end

    # @api private
    def xlsx_blob
      XyWorkbookWriter.new(self).blob
    end

    # @api private
    # Whether the worksheet carries a third column of bubble sizes.
    def bubble?
      false
    end

    def inspect
      "#<#{self.class.name} series=#{size}>"
    end

    private

    def series_class
      XySeries
    end
  end

  # Data for a bubble chart: series of (x, y, size) points.
  class BubbleChartData < XyChartData
    # @api private
    def bubble_sizes_ref(series)
      column_ref(series, "C")
    end

    # @api private
    def bubble?
      true
    end

    private

    def series_class
      BubbleSeries
    end
  end

  # One series of a scatter chart.
  class XySeries
    include Enumerable

    attr_reader :chart_data, :index, :name, :number_format, :points

    def initialize(chart_data, index, name, number_format)
      @chart_data = chart_data
      @index = index
      @name = name
      @number_format = number_format
      @points = []
    end

    # @return [self] so points can be chained
    def add_point(x, y)
      @points << [x, y]
      self
    end

    def each(&)
      @points.each(&)
    end

    def size
      @points.size
    end

    def x_values
      @points.map(&:first)
    end

    def y_values
      @points.map { |point| point[1] }
    end

    def name_ref
      @chart_data.series_name_ref(self)
    end

    def x_values_ref
      @chart_data.x_values_ref(self)
    end

    def y_values_ref
      @chart_data.y_values_ref(self)
    end

    def inspect
      "#<#{self.class.name} #{@name.inspect} #{size} points>"
    end
  end

  # One series of a bubble chart, whose points carry a size as well.
  class BubbleSeries < XySeries
    def add_point(x, y, size)
      @points << [x, y, size]
      self
    end

    def bubble_sizes
      @points.map { |point| point[2] }
    end

    def bubble_sizes_ref
      @chart_data.bubble_sizes_ref(self)
    end
  end

  # Writes the workbook behind a scatter or bubble chart.
  #
  # The layout differs from a category chart: each series gets its own block of
  # rows, since the x values belong to the series rather than being shared.
  class XyWorkbookWriter < ChartWorkbookWriter
    private

    def rows_xml
      @chart_data.series.flat_map { |series| series_rows(series) }.join
    end

    def series_rows(series)
      title_row = @chart_data.title_row(series)
      rows = [title_row_xml(series, title_row)]
      series.points.each_with_index do |point, offset|
        rows << point_row_xml(point, title_row + 1 + offset)
      end
      rows
    end

    # Column A is left empty on the title row; the series name sits in B, and
    # a bubble chart labels its third column.
    def title_row_xml(series, row)
      cells = [string_cell("B#{row}", series.name)]
      cells << string_cell("C#{row}", "Size") if @chart_data.bubble?
      %(<row r="#{row}">#{cells.join}</row>)
    end

    def point_row_xml(point, row)
      cells = point.each_with_index.map do |value, column|
        number_cell("#{(65 + column).chr}#{row}", value)
      end
      %(<row r="#{row}">#{cells.join}</row>)
    end
  end
end
