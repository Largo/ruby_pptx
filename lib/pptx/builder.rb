# frozen_string_literal: true

module Pptx
  # Build a presentation declaratively.
  #
  # This is a thin layer over the ordinary API — every method here is a call
  # the imperative API can make directly — but it keeps position and size out
  # of positional arguments and lets a whole deck read top to bottom:
  #
  #   deck = Pptx.build do |d|
  #     d.slide("Title Slide") do |s|
  #       s.title = "Annual Report"
  #       s.subtitle = "Prepared in Ruby"
  #     end
  #
  #     d.section("Detail") do
  #       d.slide("Blank") do |s|
  #         s.shape :ROUNDED_RECTANGLE, at: [Pptx.inches(1), Pptx.inches(1)],
  #                                     size: [Pptx.inches(3), Pptx.inches(1)],
  #                                     fill: "1F497D", text: "Next steps"
  #       end
  #     end
  #   end
  #
  #   deck.save("out.pptx")
  #
  # @param template [String, IO, nil] a deck to start from; the built-in
  #   default template when omitted
  # @return [Presentation]
  def self.build(template = nil)
    presentation = template ? Presentation.open(template) : Presentation.new_default
    yield DeckBuilder.new(presentation) if block_given?
    presentation
  end

  # The object yielded by {Pptx.build}.
  class DeckBuilder
    attr_reader :presentation

    def initialize(presentation)
      @presentation = presentation
      @current_section = nil
    end

    # Add a slide based on +layout+, which may be a layout name, an index, or
    # a {SlideLayout}.
    #
    # @return [Slide]
    def slide(layout = "Title and Content")
      slide_layout = resolve_layout(layout)
      slide = @presentation.slides.add(slide_layout)
      @current_section << slide if @current_section
      yield SlideBuilder.new(slide) if block_given?
      slide
    end

    # Group the slides added inside the block into a named section.
    #
    # @return [Section]
    def section(name)
      previous = @current_section
      @current_section = @presentation.sections.add(name)
      begin
        yield if block_given?
        @current_section
      ensure
        @current_section = previous
      end
    end

    # Set the slide size, e.g. `deck.slide_size = :widescreen`.
    #
    # Accepts `:widescreen` (13.333 x 7.5in), `:standard` (10 x 7.5in), or an
    # explicit [width, height] pair.
    def slide_size=(value)
      width, height =
        case value
        when :widescreen then [Pptx.inches(13.333), Pptx.inches(7.5)]
        when :standard then [Pptx.inches(10), Pptx.inches(7.5)]
        else value
        end
      @presentation.slide_width = width
      @presentation.slide_height = height
      value
    end

    private

    def resolve_layout(layout)
      return layout if layout.is_a?(SlideLayout)

      @presentation.slide_layouts[layout] ||
        raise(NotFoundError, "no slide layout #{layout.inspect}")
    end
  end

  # The object yielded by {DeckBuilder#slide}.
  class SlideBuilder
    attr_reader :slide

    def initialize(slide)
      @slide = slide
    end

    # The title placeholder's text.
    def title=(text)
      placeholder = @slide.shapes.title or
        raise NotFoundError, "this slide's layout has no title placeholder"

      placeholder.text = text
      text
    end

    # The text of the placeholder with `idx` 1, which is the body on a content
    # layout and the subtitle on a title layout.
    def body=(text)
      self[1] = text
    end
    alias subtitle= body=

    # Set the text of the placeholder with the given `idx`.
    def []=(idx, text)
      placeholder = @slide.placeholders[idx] or
        raise NotFoundError, "this slide has no placeholder with idx #{idx}"

      placeholder.text = text
      text
    end

    # Add a text box.
    #
    # @return [Shape]
    def text(content, at:, size: nil, font_size: nil, bold: nil, italic: nil,
             color: nil, font: nil, align: nil)
      left, top = at
      width, height = size || [Pptx.inches(4), Pptx.inches(1)]
      box = @slide.shapes.add_textbox(left, top, width, height)
      box.text_frame.text = content
      style_text(box.text_frame, font_size: font_size, bold: bold, italic: italic,
                                 color: color, font: font, align: align)
      box
    end

    # Add an auto shape.
    #
    # @return [Shape]
    def shape(shape_type, at:, size:, fill: nil, line: nil, text: nil, **text_options)
      left, top = at
      width, height = size
      auto_shape = @slide.shapes.add_shape(shape_type, left, top, width, height)
      apply_fill(auto_shape, fill)
      apply_line(auto_shape, line)
      if text
        auto_shape.text_frame.text = text
        style_text(auto_shape.text_frame, **text_options)
      end
      auto_shape
    end

    # Add a picture. Omit +size+ for the image's native size, or give one
    # dimension to scale the other with it.
    #
    # @return [Picture]
    def picture(image, at:, width: nil, height: nil)
      left, top = at
      @slide.shapes.add_picture(image, left, top, width: width, height: height)
    end

    # Add a table filled from +rows+, an array of arrays.
    #
    # @return [Table]
    def table(rows, at:, size:, header: true)
      rows = rows.to_a
      raise ArgumentError, "table needs at least one row" if rows.empty?

      left, top = at
      width, height = size
      columns = rows.map(&:size).max
      frame = @slide.shapes.add_table(rows.size, columns, left, top, width, height)
      table = frame.table
      table.first_row = header
      rows.each_with_index do |row, row_index|
        row.each_with_index { |value, col| table.cell(row_index, col).text = value.to_s }
      end
      table
    end

    # Add a chart.
    #
    #   s.chart :COLUMN_CLUSTERED, categories: %w[East West],
    #                              series: { "Q1" => [1, 2] },
    #                              at: [x, y], size: [w, h]
    #
    # @return [Chart]
    def chart(chart_type, categories:, series:, at:, size:)
      data = ChartData.new
      data.categories = categories
      series.each { |name, values| data.add_series(name, values) }
      left, top = at
      width, height = size
      @slide.shapes.add_chart(chart_type, left, top, width, height, data).chart
    end

    private

    def style_text(text_frame, font_size: nil, bold: nil, italic: nil,
                   color: nil, font: nil, align: nil)
      paragraph = text_frame.paragraphs.first
      paragraph.alignment = align if align
      # Apply to every run so a multi-run value styles consistently.
      runs = text_frame.paragraphs.flat_map(&:runs)
      runs.each do |run|
        run.font.size = font_size if font_size
        run.font.bold = bold unless bold.nil?
        run.font.italic = italic unless italic.nil?
        run.font.name = font if font
        run.font.color.rgb = rgb(color) if color
      end
    end

    def apply_fill(shape, fill)
      return if fill.nil?
      return shape.fill.background if fill == :none

      shape.fill.solid
      shape.fill.fore_color.rgb = rgb(fill)
    end

    def apply_line(shape, line)
      return if line.nil?

      color, width = line.is_a?(Hash) ? [line[:color], line[:width]] : [line, nil]
      shape.line.color.rgb = rgb(color) if color
      shape.line.width = width if width
    end

    # A colour may be given as an RGBColor or as a hex string.
    def rgb(value) = value.is_a?(RGBColor) ? value : RGBColor.from_string(value)
  end
end
