# frozen_string_literal: true

require "pptx/errors"

module Pptx
  # The categories and series a chart depicts.
  #
  #   data = Pptx::ChartData.new
  #   data.categories = ["East", "West", "Midwest"]
  #   data.add_series("Q1", [1.2, 2.0, 3.5])
  #   data.add_series("Q2", [4.1, 5.0, 6.2])
  #
  # The values also become an embedded Excel workbook, which is what PowerPoint
  # opens when the user chooses "Edit Data". The worksheet references in the
  # chart XML are derived from the same layout, so the two always agree.
  class ChartData
    include Enumerable

    # The row the first data value sits on; row 1 holds the series names.
    FIRST_DATA_ROW = 2
    # Column A holds the category labels, so series start at column B.
    FIRST_SERIES_COLUMN = 2
    WORKSHEET_NAME = "Sheet1"
    DEFAULT_NUMBER_FORMAT = "General"

    attr_reader :series, :number_format, :categories

    def initialize(number_format: DEFAULT_NUMBER_FORMAT)
      @categories = []
      @series = []
      @number_format = number_format
    end

    def categories=(values)
      @categories = values.to_a
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

    def each(&) = @series.each(&)

    def size = @series.size

    def category_count = @categories.size

    # True when every category is a number, which makes the category axis
    # numeric rather than textual.
    def numeric_categories?
      !@categories.empty? && @categories.all?(Numeric)
    end

    # @api private
    # Categories are always column A, one row per category.
    def categories_ref
      last_row = FIRST_DATA_ROW + category_count - 1
      "#{WORKSHEET_NAME}!$A$#{FIRST_DATA_ROW}:$A$#{last_row}"
    end

    # @api private
    def series_name_ref(series) = "#{WORKSHEET_NAME}!$#{column_letter(series)}$1"

    # @api private
    def series_values_ref(series)
      letter = column_letter(series)
      last_row = FIRST_DATA_ROW + category_count - 1
      "#{WORKSHEET_NAME}!$#{letter}$#{FIRST_DATA_ROW}:$#{letter}$#{last_row}"
    end

    # @api private
    # The Excel column letter for a series, e.g. the third series is "D".
    def column_letter(series) = ChartData.column_reference(FIRST_SERIES_COLUMN + series.index)

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
    def xlsx_blob = ChartWorkbookWriter.new(self).blob

    def inspect = "#<Pptx::ChartData categories=#{category_count} series=#{size}>"
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

    def each(&) = @values.each(&)

    def size = @values.size

    def categories = @chart_data.categories

    def name_ref = @chart_data.series_name_ref(self)

    def values_ref = @chart_data.series_values_ref(self)

    def categories_ref = @chart_data.categories_ref

    def inspect = "#<Pptx::ChartSeries #{@name.inspect} #{size} values>"
  end
end
