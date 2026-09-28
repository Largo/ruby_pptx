# frozen_string_literal: true

require "ruby_pptx/errors"
require "ruby_pptx/chart/categories"

module Pptx
  # The categories and series a chart depicts.
  #
  #   data = Pptx::ChartData.new
  #   data.categories = ["East", "West", "Midwest"]
  #   data.add_series("Q1", [1.2, 2.0, 3.5])
  #   data.add_series("Q2", [4.1, 5.0, 6.2])
  #
  # Categories can be grouped, which draws a multi-level axis, and can be
  # dates, which draws a date axis; see {ChartDataCategories}.
  #
  # The values also become an embedded Excel workbook, which is what PowerPoint
  # opens when the user chooses "Edit Data". The worksheet references in the
  # chart XML are derived from the same layout, so the two always agree.
  class ChartData
    include Enumerable

    # The row the first data value sits on; row 1 holds the series names.
    FIRST_DATA_ROW = 2
    # Category labels fill the first columns, one per level; series follow.
    WORKSHEET_NAME = "Sheet1"
    DEFAULT_NUMBER_FORMAT = "General"

    attr_reader :series, :number_format, :categories

    def initialize(number_format: DEFAULT_NUMBER_FORMAT)
      @categories = ChartDataCategories.new
      @series = []
      @number_format = number_format
    end

    # Replace the categories: a list of labels, or a Hash of groups whose
    # values are their sub-categories.
    def categories=(values)
      @categories = values.is_a?(ChartDataCategories) ? values : ChartDataCategories.from(values)
    end

    # Add a series of values, one per category.
    #
    # @return [ChartSeries]
    def add_series(name, values, number_format: nil)
      series = ChartSeries.new(self, @series.size, name, values.to_a,
                               number_format || @number_format)
      @series << series
      series
    end

    def each(&)
      @series.each(&)
    end

    def size
      @series.size
    end

    # The number of categories plotted: the leaves, when they are grouped.
    def category_count
      @categories.leaf_count
    end

    # True when the categories are numbers or dates, which makes the category
    # cache numeric rather than textual. Decided by the first category, as
    # python-pptx decides it.
    def numeric_categories?
      @categories.numeric?
    end

    # @api private
    # The category block: one column per level, one row per leaf.
    def categories_ref
      raise Error, "chart data contains no categories" if @categories.depth.zero?

      right = ChartData.column_reference(@categories.depth)
      "#{WORKSHEET_NAME}!$A$#{FIRST_DATA_ROW}:$#{right}$#{FIRST_DATA_ROW + category_count - 1}"
    end

    # @api private
    def series_name_ref(series)
      "#{WORKSHEET_NAME}!$#{column_letter(series)}$1"
    end

    # @api private
    # As long as the series itself, which python-pptx also measures by the
    # series rather than by the categories.
    def series_values_ref(series)
      letter = column_letter(series)
      last_row = FIRST_DATA_ROW + series.size - 1
      "#{WORKSHEET_NAME}!$#{letter}$#{FIRST_DATA_ROW}:$#{letter}$#{last_row}"
    end

    # @api private
    # The Excel column letter for a series: after the category columns.
    def column_letter(series)
      ChartData.column_reference(1 + @categories.depth + series.index)
    end

    # Excel's bijective base-26 column names: 1 => "A", 27 => "AA".
    def self.column_reference(column_number)
      raise ArgumentError, "column number must be positive" unless column_number.positive?

      letters = +""
      remaining = column_number
      while remaining.positive?
        remaining, remainder = (remaining - 1).divmod(26)
        letters.prepend((remainder + 65).chr)
      end
      letters
    end

    # @api private
    # Bytes of the embedded Excel workbook holding this data.
    def xlsx_blob
      ChartWorkbookWriter.new(self).blob
    end

    def inspect
      "#<Pptx::ChartData categories=#{category_count} series=#{size}>"
    end
  end

  # One series of a chart: a name and one value per category.
  class ChartSeries
    include Enumerable

    attr_reader :chart_data, :index, :name, :values, :number_format

    def initialize(chart_data, index, name, values, number_format)
      @chart_data = chart_data
      @index = index
      @name = name
      @values = values
      @number_format = number_format
    end

    def each(&)
      @values.each(&)
    end

    def size
      @values.size
    end

    # Append one value, for building a series up point by point.
    #
    # python-pptx's add_data_point also takes a per-point number format, but
    # nothing it writes ever reads it; the series' format is the one used.
    def <<(value)
      @values << value
      self
    end

    def categories
      @chart_data.categories
    end

    def name_ref
      @chart_data.series_name_ref(self)
    end

    def values_ref
      @chart_data.series_values_ref(self)
    end

    def categories_ref
      @chart_data.categories_ref
    end

    def inspect
      "#<Pptx::ChartSeries #{@name.inspect} #{size} values>"
    end
  end
end
