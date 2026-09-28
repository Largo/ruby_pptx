# frozen_string_literal: true

require "date"

module Pptx
  # The categories of a chart being built -- a plain list, or a tree when
  # categories are grouped.
  #
  #   data.categories = %w[East West Mid]
  #
  #   data.categories = { "East" => %w[North South], "West" => %w[Coast] }
  #
  #   east = data.categories.add_category("East")
  #   east.add_sub_category("North")
  #
  # Grouped categories become a multi-level axis: the leaves are plotted,
  # and each level above labels a run of them. Every branch must be equally
  # deep.
  #
  # Dates make a date axis, and numbers a numeric one; both are decided by the
  # first category, as python-pptx decides them.
  class ChartDataCategories
    include Enumerable

    GENERAL = "General"
    DATE_FORMAT = "yyyy\\-mm\\-dd"

    # Excel's own number format for dates, used unless one is set.
    attr_writer :number_format

    def initialize
      @categories = []
      @number_format = nil
    end

    # Build from a list of labels, or from a Hash whose values are the
    # sub-categories of each key, nested as deep as needed.
    def self.from(value)
      new.tap { |categories| categories.add_all(value) }
    end

    # @api private
    def add_all(value)
      if value.is_a?(Hash)
        value.each { |label, children| add_category(label).add_all(children) }
      else
        Array(value).each { |label| add_category(label) }
      end
      self
    end

    # @return [ChartDataCategory] the new top-level category
    def add_category(label)
      ChartDataCategory.new(label, self).tap { |category| @categories << category }
    end

    def each(&) = @categories.each(&)

    def [](index) = @categories[index]

    def size = @categories.size
    alias length size

    def empty? = @categories.empty?

    # Levels of labels: 0 with no categories, 1 for a plain list.
    #
    # @raise [Error] when the branches are not all equally deep
    def depth
      return 0 if @categories.empty?

      depths = @categories.map(&:depth).uniq
      raise Error, "category depth not uniform" if depths.size > 1

      depths.first
    end

    # The number of categories actually plotted.
    def leaf_count = @categories.sum(&:leaf_count)

    # Each level as [index, label] pairs, leaf level first. The index is the
    # leaf position where that label's run begins.
    def levels
      collect_levels(@categories)
    end

    # The leaf labels in plotting order.
    def leaf_labels = levels.first.to_a.map(&:last)

    def dates? = depth == 1 && date?(first.label)

    def numeric? = depth == 1 && (first.label.is_a?(Numeric) || date?(first.label))

    # The number format for the categories: as set, or a date format for
    # dates, or "General".
    def number_format
      return @number_format unless @number_format.nil?

      dates? ? DATE_FORMAT : GENERAL
    end

    # @api private
    # The leaf offset at which +category+'s run begins.
    def index_of(category)
      offset = 0
      @categories.each do |candidate|
        return offset if candidate.equal?(category)

        offset += candidate.leaf_count
      end
      raise Error, "category not in these categories"
    end

    def inspect = "#<Pptx::ChartDataCategories depth=#{depth} #{leaf_labels.inspect}>"

    private

    def collect_levels(categories)
      children = categories.flat_map(&:sub_categories)
      lower = children.empty? ? [] : collect_levels(children)
      lower + [categories.map { |category| [category.idx, category.label] }]
    end

    def date?(value) = value.is_a?(Date) || value.is_a?(Time)
  end

  # One category, or a group of them when it has sub-categories.
  class ChartDataCategory
    attr_reader :sub_categories

    def initialize(label, parent)
      @label = label
      @parent = parent
      @sub_categories = []
    end

    # The label; "" for a category given as nil.
    def label = @label.nil? ? "" : @label

    # @return [ChartDataCategory] the new sub-category
    def add_sub_category(label)
      ChartDataCategory.new(label, self).tap { |category| @sub_categories << category }
    end

    # @api private
    def add_all(value)
      if value.is_a?(Hash)
        value.each { |label, children| add_sub_category(label).add_all(children) }
      else
        Array(value).each { |label| add_sub_category(label) }
      end
      self
    end

    def depth
      return 1 if @sub_categories.empty?

      depths = @sub_categories.map(&:depth).uniq
      raise Error, "category depth not uniform" if depths.size > 1

      depths.first + 1
    end

    def leaf_count = @sub_categories.empty? ? 1 : @sub_categories.sum(&:leaf_count)

    # The leaf position at which this category's run begins.
    def idx = @parent.index_of(self)

    # @api private
    def index_of(sub_category)
      offset = idx
      @sub_categories.each do |candidate|
        return offset if candidate.equal?(sub_category)

        offset += candidate.leaf_count
      end
      raise Error, "sub-category not in this category"
    end

    # The label as a numeric category's cached value: a date as its Excel
    # serial day number with one decimal, as python-pptx writes it.
    def numeric_str_val(date_1904: false)
      value = @label
      if value.is_a?(Date) || value.is_a?(Time)
        return format("%.1f",
                      excel_day_number(value.to_date, date_1904))
      end

      value.to_s
    end

    # The label as it appears in the category cache and worksheet.
    def to_s = label.to_s

    def inspect = "#<Pptx::ChartDataCategory #{label.inspect}>"

    private

    # Days since Excel's epoch. Excel treats 1900 as a leap year, so every
    # date after 28 February 1900 is one day further on than the calendar
    # says.
    def excel_day_number(date, date_1904)
      epoch = date_1904 ? Date.new(1904, 1, 1) : Date.new(1899, 12, 31)
      days = (date - epoch).to_i
      !date_1904 && days > 59 ? days + 1 : days
    end
  end
end
