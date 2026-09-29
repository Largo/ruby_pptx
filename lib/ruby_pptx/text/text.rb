# frozen_string_literal: true

require "ruby_pptx/pattern_matching"

require "ruby_pptx/element_proxy"
require "ruby_pptx/dml/color"
require "ruby_pptx/dml/fill"
require "ruby_pptx/oxml/text"
require "ruby_pptx/enum/text"
require "ruby_pptx/action"
require "ruby_pptx/text/fitter"

module Pptx
  # The text inside a shape.
  #
  # A text frame always holds at least one paragraph. Reading {#text} joins
  # the paragraphs with "\n"; assigning it replaces everything.
  #
  #   frame = shape.text_frame
  #   frame.text = "First line\nSecond line"
  #   frame.paragraphs.first.runs.first.font.bold = true
  class TextFrame
    include PatternMatching

    pattern_keys :text, :paragraphs, :word_wrap, :auto_size, :vertical_anchor

    attr_reader :parent, :element

    def initialize(tx_body, parent)
      @element = tx_body
      @parent = parent
    end

    def part
      @parent.part
    end

    # The paragraphs of this text frame; there is always at least one.
    #
    # @return [Array<Paragraph>]
    def paragraphs
      @element.p_list.map { |p| Paragraph.new(p, self) }
    end

    # Append an empty paragraph and return it.
    def add_paragraph
      Paragraph.new(@element.add_p, self)
    end

    # All the text, with "\n" between paragraphs and "\v" for each soft line
    # break.
    def text
      paragraphs.map(&:text).join("\n")
    end

    # Replace all text. Each "\n" starts a new paragraph, each "\v" a soft
    # line break within one.
    def text=(value)
      @element.clear_content
      value.to_s.split("\n", -1).each do |paragraph_text|
        @element.add_p.append_text(paragraph_text)
      end
      @element.unclear_content
    end

    # Remove all text, leaving a single empty paragraph.
    def clear
      @element.clear_content
      @element.unclear_content
      self
    end

    def margin_left
      body_properties.lIns
    end

    def margin_left=(value)
      body_properties.lIns = value
    end

    def margin_right
      body_properties.rIns
    end

    def margin_right=(value)
      body_properties.rIns = value
    end

    def margin_top
      body_properties.tIns
    end

    def margin_top=(value)
      body_properties.tIns = value
    end

    def margin_bottom
      body_properties.bIns
    end

    def margin_bottom=(value)
      body_properties.bIns = value
    end

    # @return [Pptx::Enum::MSO_ANCHOR, nil] nil when inherited
    def vertical_anchor
      body_properties.anchor
    end

    def vertical_anchor=(value)
      body_properties.anchor = value
    end

    # @return [Pptx::Enum::MSO_AUTO_SIZE, nil] nil when inherited
    def auto_size
      body_properties.autofit
    end

    def auto_size=(value)
      body_properties.autofit = value
    end

    # True, false, or nil when the setting is inherited.
    def word_wrap
      case body_properties.wrap
      when Oxml::SimpleTypes::ST_TextWrappingType::SQUARE then true
      when Oxml::SimpleTypes::ST_TextWrappingType::NONE then false
      end
    end

    def word_wrap=(value)
      unless [true, false, nil].include?(value)
        raise ArgumentError, "word_wrap must be true, false or nil, got #{value.inspect}"
      end

      body_properties.wrap =
        case value
        when true then Oxml::SimpleTypes::ST_TextWrappingType::SQUARE
        when false then Oxml::SimpleTypes::ST_TextWrappingType::NONE
        end
    end

    # Shrink the text until it fits the shape, and apply that size to all of it.
    #
    # Turns word wrap on and autofit off, then sets every run -- and the
    # end-paragraph properties, so typing in PowerPoint continues in the same
    # font -- to +font_family+ at the size found.
    #
    #   frame.fit_text(font_file: "/usr/share/fonts/.../DejaVuSans.ttf",
    #                  font_family: "DejaVu Sans", max_size: 28)
    #
    # A +font_file+ is required: measuring needs the actual glyph outlines, and
    # this gem does not go looking through system font directories for them.
    # The size is measured from the font's own metrics rather than by
    # rendering, so it can differ from python-pptx's by a point on text that
    # only just fits; see PORTING.md.
    #
    # @return [Integer] the point size applied
    def fit_text(font_file:, font_family: "Calibri", max_size: 18, bold: false, italic: false)
      raise Error, "cannot fit text in a text frame with no text" if text.empty?

      size = TextFitter.best_fit_font_size(text, extents: extents, max_size: max_size,
                                                 font_file: font_file)
      raise Error, "text does not fit at any size up to #{max_size}pt" if size.nil?

      apply_fit(font_family, size, bold, italic)
      size
    end

    # The area text actually gets to occupy: the shape less its margins.
    #
    # @return [Array(Integer, Integer)] width and height in EMU
    def extents
      [@parent.width - margin_left - margin_right, @parent.height - margin_top - margin_bottom]
    end

    def inspect
      "#<Pptx::TextFrame #{text.inspect}>"
    end

    private

    def apply_fit(family, size, bold, italic)
      self.auto_size = Enum::MSO_AUTO_SIZE::NONE
      self.word_wrap = true
      each_character_properties do |properties|
        font = Font.new(properties)
        font.name = family
        font.size = Pptx.pt(size)
        font.bold = bold
        font.italic = italic
      end
    end

    # Every run in the frame, plus each paragraph's end-paragraph properties.
    def each_character_properties
      @element.p_list.each do |p|
        p.content_children.each { |child| yield child.get_or_add_rPr }
        yield p.get_or_add_endParaRPr
      end
    end

    def body_properties
      @element.bodyPr
    end
  end

  # One paragraph of a text frame.
  class Paragraph
    include PatternMatching

    pattern_keys :text, :runs, :level, :alignment

    attr_reader :element, :parent

    def initialize(p, parent)
      @element = p
      @parent = parent
    end

    def part
      @parent.part
    end

    # @return [Array<Run>]
    def runs
      @element.r_list.map { |r| Run.new(r, self) }
    end

    def add_run(text = nil)
      Run.new(@element.add_run(text), self)
    end

    def add_line_break
      @element.add_line_break
      self
    end

    # The text of this paragraph, with "\v" for each soft line break.
    def text
      @element.text
    end

    # Replace this paragraph's text, turning "\v" (or "\n") into line breaks
    # rather than new paragraphs -- a paragraph cannot contain another.
    def text=(value)
      clear
      @element.append_text(value)
    end

    # Remove the content, keeping the paragraph and its properties.
    def clear
      @element.content_children.each { |child| @element.remove(child) }
      self
    end

    # @return [Pptx::Enum::PP_ALIGN, nil] nil when inherited
    def alignment
      paragraph_properties.algn
    end

    def alignment=(value)
      paragraph_properties.algn = value
    end

    # Outline level, 0 for the top level.
    def level
      paragraph_properties.lvl
    end

    def level=(value)
      paragraph_properties.lvl = value
    end

    # A Float is a number of lines; a {Pptx::Length} is a fixed distance.
    def line_spacing
      paragraph_properties.line_spacing
    end

    def line_spacing=(value)
      paragraph_properties.line_spacing = value
    end

    def space_before
      paragraph_properties.space_before
    end

    def space_before=(value)
      paragraph_properties.space_before = value
    end

    def space_after
      paragraph_properties.space_after
    end

    def space_after=(value)
      paragraph_properties.space_after = value
    end

    # The bullet: a String such as "•" for a character bullet, :none to
    # suppress the one the master would draw, an autonumber scheme such as
    # :arabicPeriod ("1.") or :alphaLcParenR ("a)"), or nil to inherit.
    def bullet
      paragraph_properties.bullet
    end

    def bullet=(value)
      paragraph_properties.bullet = value
    end

    # Distance from the left edge of the text box to the text, as a
    # {Pptx::Length}; nil when inherited.
    def margin_left
      paragraph_properties.marL
    end

    def margin_left=(value)
      paragraph_properties.marL = value
    end

    # Offset of the first line from {#margin_left}. Negative gives a hanging
    # indent, where the bullet sits in the margin: margin_left 0.25", indent
    # -0.25". nil when inherited.
    def indent
      paragraph_properties.indent
    end

    def indent=(value)
      paragraph_properties.indent = value
    end

    # The default character formatting for runs in this paragraph.
    def font
      Font.new(paragraph_properties.get_or_add_defRPr)
    end

    def inspect
      "#<Pptx::Paragraph #{text.inspect}>"
    end

    private

    def paragraph_properties
      @element.get_or_add_pPr
    end
  end

  # A run: a span of text sharing one set of character properties.
  class Run
    include PatternMatching

    pattern_keys :text, :font, :hyperlink

    attr_reader :element, :parent

    def initialize(r, parent)
      @element = r
      @parent = parent
    end

    def part
      @parent.part
    end

    def text
      @element.text
    end

    def text=(value)
      @element.text = value
    end

    def font
      Font.new(@element.get_or_add_rPr)
    end

    # What happens when this run of text is clicked.
    def click_action
      @click_action ||= ActionSetting.new(@element.get_or_add_rPr, self)
    end

    # The URL this run links to, or nil.
    def hyperlink
      click_action.url
    end

    def hyperlink=(url)
      click_action.address = url
    end

    def inspect
      "#<Pptx::Run #{text.inspect}>"
    end
  end

  # Character formatting: typeface, size, weight, colour.
  #
  # Every property returns nil when the value is inherited from the style
  # hierarchy rather than set here, and assigning nil restores that
  # inheritance.
  class Font
    include PatternMatching

    pattern_keys :name, :size, :bold, :italic, :underline

    attr_reader :element

    def initialize(r_pr)
      @element = r_pr
    end

    def bold
      @element.b
    end

    def bold=(value)
      @element.b = value
    end

    def italic
      @element.i
    end

    def italic=(value)
      @element.i = value
    end

    # The typeface name, or nil when inherited from the theme.
    def name
      @element.latin&.typeface
    end

    def name=(value)
      if value.nil?
        @element.remove_latin
      else
        @element.get_or_add_latin.typeface = value
      end
    end

    # @return [Pptx::Length, nil]
    def size
      centipoints = @element.sz
      centipoints && Pptx::Length.centipoints(centipoints)
    end

    def size=(value)
      @element.sz = value.nil? ? nil : Pptx::Length.coerce(value).centipoints
    end

    # true for a single underline, false for none, a
    # {Pptx::Enum::MSO_UNDERLINE} member for anything fancier, nil when
    # inherited.
    def underline
      value = @element.u
      return nil if value.nil?
      return false if value == Enum::MSO_UNDERLINE::NONE
      return true if value == Enum::MSO_UNDERLINE::SINGLE_LINE

      value
    end

    def underline=(value)
      @element.u =
        case value
        when true then Enum::MSO_UNDERLINE::SINGLE_LINE
        when false then Enum::MSO_UNDERLINE::NONE
        else value
        end
    end

    # @return [Pptx::Enum::MSO_LANGUAGE_ID, nil]
    def language
      @element.lang
    end

    def language=(value)
      @element.lang = value
    end

    # The fill of the text itself.
    def fill
      @fill ||= FillFormat.from_fill_parent(@element)
    end

    # The text colour. Accessing it makes the fill solid, since a colour has
    # to live on some fill.
    def color
      fill.solid if fill.type.nil?
      fill.fore_color
    end

    def inspect
      "#<Pptx::Font name=#{name.inspect} size=#{size&.pt} bold=#{bold.inspect}>"
    end
  end
end
