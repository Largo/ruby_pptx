# frozen_string_literal: true

require "ruby_pptx/text/text"

module Pptx
  # One outline level of a list style: the formatting a paragraph at that
  # level gets unless it says otherwise.
  #
  #   master.text_styles.body.level(0).tap do |l1|
  #     l1.bullet = "•"
  #     l1.margin_left = Pptx.inches(0.5)
  #     l1.indent = -Pptx.inches(0.5)
  #     l1.font.size = Pptx.pt(37)
  #   end
  class ParagraphStyle
    include ParagraphFormatting

    attr_reader :element

    def initialize(ppr)
      @element = ppr
    end

    def inspect
      "#<Pptx::ParagraphStyle #{@element.nsptag}>"
    end

    private

    def paragraph_properties
      @element
    end
  end

  # A list style: a {ParagraphStyle} per outline level, 0 to 8.
  class ListStyle
    attr_reader :element

    def initialize(element)
      @element = element
    end

    # The style for outline level +level+, counted from 0 like
    # {Paragraph#level}. Created empty -- inheriting everything -- if the
    # style had none.
    #
    # @return [ParagraphStyle]
    def level(level)
      ParagraphStyle.new(@element.get_or_add_level(level))
    end

    # The properties that apply at every level before the level's own.
    def default
      ParagraphStyle.new(@element.get_or_add_defPPr)
    end

    def inspect
      "#<Pptx::ListStyle #{@element.nsptag}>"
    end
  end

  # A master's text styles: one for titles, one for body placeholders and
  # one for everything else (text boxes, shapes, tables).
  class TextStyles
    attr_reader :element

    def initialize(tx_styles)
      @element = tx_styles
    end

    # @return [ListStyle]
    def title
      ListStyle.new(@element.get_or_add_titleStyle)
    end

    def body
      ListStyle.new(@element.get_or_add_bodyStyle)
    end

    def other
      ListStyle.new(@element.get_or_add_otherStyle)
    end

    def inspect
      "#<Pptx::TextStyles>"
    end
  end
end
