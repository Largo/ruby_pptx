# frozen_string_literal: true

require "ruby_pptx/element_proxy"
require "ruby_pptx/dml/color"
require "ruby_pptx/oxml/theme"

module Pptx
  # The colours and fonts a slide master draws from.
  #
  #   theme = master.theme
  #   theme.name = "Corporate"
  #   theme.colors[:accent1].rgb = Pptx::RGBColor.from_string("1F497D")
  #   theme.fonts.major = "Georgia"
  #
  # A theme also carries a *format* scheme -- the numbered fill, line and
  # effect styles that shapes and slide backgrounds refer to by `idx`. Those
  # are inherited from the base theme and not editable here; writing a
  # coherent set by hand is a design exercise, not a configuration one.
  class Theme < PartElementProxy
    def name = @element.name

    def name=(value)
      # PowerPoint shows the theme name in three places and expects them to
      # agree, so the schemes are renamed with it.
      @element.name = value
      @element.clrScheme.name = value
      @element.fontScheme.name = value
    end

    # @return [ThemeColors]
    def colors = @colors ||= ThemeColors.new(@element.clrScheme, self)

    # @return [ThemeFonts]
    def fonts = @fonts ||= ThemeFonts.new(@element.fontScheme, self)

    def inspect = "#<Pptx::Theme #{name.inspect}>"
  end

  # The twelve theme colours, addressed by name.
  #
  #   theme.colors[:accent1].rgb = Pptx::RGBColor.from_string("C0504D")
  #   theme.colors[:accent1].rgb   #=> #<Pptx::RGBColor C0504D>
  #
  # Each entry is an ordinary {Pptx::ColorFormat}, so it takes an explicit RGB
  # value the same way any other colour in the library does.
  class ThemeColors < ParentedElementProxy
    include Enumerable

    # The names, in the order PowerPoint lists them.
    NAMES = Oxml::CT_ColorScheme::SLOTS.map(&:to_sym).freeze

    # @return [Pptx::ColorFormat]
    def [](name) = ColorFormat.from_color_choice_parent(@element.slot(name))

    def each
      return enum_for(:each) { NAMES.size } unless block_given?

      NAMES.each { |name| yield name, self[name] }
      self
    end

    # Set several at once, which is how a palette is usually handed over.
    #
    #   theme.colors.update(accent1: "1F497D", accent2: "C0504D")
    #
    # Values may be {Pptx::RGBColor} or a hex string.
    def update(**colors)
      colors.each do |name, value|
        self[name].rgb = value.is_a?(RGBColor) ? value : RGBColor.from_string(value.to_s)
      end
      self
    end

    def names = NAMES

    def inspect = "#<Pptx::ThemeColors #{NAMES.size} colours>"
  end

  # The heading and body typefaces of a theme.
  #
  # PowerPoint calls these "major" and "minor"; they are what `+mj-lt` and
  # `+mn-lt` resolve to in placeholder text.
  class ThemeFonts < ParentedElementProxy
    def major = @element.majorFont.typeface

    def major=(value)
      @element.majorFont.typeface = value
    end

    def minor = @element.minorFont.typeface

    def minor=(value)
      @element.minorFont.typeface = value
    end

    def inspect = "#<Pptx::ThemeFonts major=#{major.inspect} minor=#{minor.inspect}>"
  end
end
