# frozen_string_literal: true

require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/oxml/simple_types"
require "pptx/oxml/dml/fill"
require "pptx/enum/text"
require "pptx/enum/lang"

module Pptx
  module Oxml
    # The fill choice and its successor list, shared by run properties and
    # shape properties.
    FILL_CHOICES = %w[a:noFill a:solidFill a:gradFill a:blipFill a:pattFill a:grpFill].freeze

    # `a:t`, the literal text of a run.
    class CT_TextCharacterText < Element
      tag "a:t"
    end

    # `a:r`, a run: a span of text sharing one set of character properties.
    class CT_RegularTextRun < Element
      tag "a:r"
      zero_or_one "a:rPr", successors: %w[a:t]
      one_and_only_one "a:t"

      def text = t.text

      def text=(value)
        t.text = CT_RegularTextRun.escape_control_chars(value.to_s)
      end

      # XML 1.0 cannot carry most control characters, so PowerPoint writes
      # them as a literal escape like "_x0007_". Tab and line-feed are legal
      # and pass through unchanged.
      def self.escape_control_chars(string)
        string.gsub(/([\x00-\x08\x0B-\x1F])/) { format("_x%04X_", ::Regexp.last_match(1).ord) }
      end
    end

    # `a:br`, a soft line break within a paragraph.
    class CT_TextLineBreak < Element
      tag "a:br"
      zero_or_one "a:rPr", successors: []

      # A break carries no text of its own; it *is* the line feed, which we
      # report as a vertical tab to match how PowerPoint puts it on the
      # clipboard.
      def text = "\v"
    end

    # `a:fld`, a slide-number or date field.
    class CT_TextField < Element
      tag "a:fld"
      zero_or_one "a:rPr", successors: %w[a:pPr a:t]
      zero_or_one "a:t", successors: []

      def text = t&.text.to_s
    end

    # `a:latin`, `a:ea`, `a:cs` and `a:sym`: a typeface reference.
    class CT_TextFont < Element
      tag "a:latin", "a:ea", "a:cs", "a:sym"
      required_attr "typeface", type: SimpleTypes::ST_TextTypeface
    end

    # `a:rPr`, `a:defRPr` and `a:endParaRPr`: character properties.
    class CT_TextCharacterProperties < Element
      tag "a:rPr", "a:defRPr", "a:endParaRPr"

      FILL_SUCCESSORS = %w[a:effectLst a:effectDag a:highlight a:uLnTx a:uLn
                           a:uFillTx a:uFill a:latin a:ea a:cs a:sym
                           a:hlinkClick a:hlinkMouseOver a:rtl a:extLst].freeze

      zero_or_one_choice FILL_CHOICES.map { |t| choice(t) },
                         successors: FILL_SUCCESSORS, as: :eg_fillProperties
      zero_or_one "a:latin",
                  successors: %w[a:ea a:cs a:sym a:hlinkClick a:hlinkMouseOver a:rtl a:extLst]
      zero_or_one "a:hlinkClick", successors: %w[a:hlinkMouseOver a:rtl a:extLst]

      optional_attr "lang", type: Enum::MSO_LANGUAGE_ID
      optional_attr "sz", type: SimpleTypes::ST_TextFontSize
      optional_attr "b", type: SimpleTypes::XsdBoolean
      optional_attr "i", type: SimpleTypes::XsdBoolean
      optional_attr "u", type: Enum::MSO_TEXT_UNDERLINE_TYPE

      def new_gradFill = CT_GradientFillProperties.new_grad_fill(self)
    end

    # `a:bodyPr`, the properties of a text body: margins, anchoring, autofit.
    class CT_TextBodyProperties < Element
      tag "a:bodyPr"

      zero_or_one_choice [choice("a:noAutofit"), choice("a:normAutofit"), choice("a:spAutoFit")],
                         successors: %w[a:scene3d a:sp3d a:flatTx a:extLst], as: :eg_textAutoFit

      optional_attr "lIns", type: SimpleTypes::ST_Coordinate32, default: Pptx::Length.emu(91_440)
      optional_attr "tIns", type: SimpleTypes::ST_Coordinate32, default: Pptx::Length.emu(45_720)
      optional_attr "rIns", type: SimpleTypes::ST_Coordinate32, default: Pptx::Length.emu(91_440)
      optional_attr "bIns", type: SimpleTypes::ST_Coordinate32, default: Pptx::Length.emu(45_720)
      optional_attr "anchor", type: Enum::MSO_VERTICAL_ANCHOR
      optional_attr "wrap", type: SimpleTypes::ST_TextWrappingType

      # @return [Pptx::Enum::MSO_AUTO_SIZE, nil] nil when inherited
      def autofit
        return Enum::MSO_AUTO_SIZE::NONE unless find("a:noAutofit").nil?
        return Enum::MSO_AUTO_SIZE::TEXT_TO_FIT_SHAPE unless find("a:normAutofit").nil?
        return Enum::MSO_AUTO_SIZE::SHAPE_TO_FIT_TEXT unless find("a:spAutoFit").nil?

        nil
      end

      def autofit=(value)
        remove_eg_textAutoFit
        return if value.nil?

        member = Enum::MSO_AUTO_SIZE.fetch(value)
        case member.name
        when :NONE then add_noAutofit
        when :TEXT_TO_FIT_SHAPE then add_normAutofit
        when :SHAPE_TO_FIT_TEXT then add_spAutoFit
        else
          raise ArgumentError, "#{member} cannot be assigned as an autofit setting"
        end
      end
    end

    # `a:normAutofit`, shrink-text-on-overflow with its font scale.
    class CT_TextNormalAutofit < Element
      tag "a:normAutofit"
      optional_attr "fontScale", type: SimpleTypes::ST_TextFontScalePercentOrPercentString,
                                 default: 100.0
    end

    # `a:spcPct`, spacing as a multiple of line height.
    class CT_TextSpacingPercent < Element
      tag "a:spcPct"
      required_attr "val", type: SimpleTypes::ST_TextSpacingPercentOrPercentString
    end

    # `a:spcPts`, spacing as a fixed distance.
    class CT_TextSpacingPoint < Element
      tag "a:spcPts"
      required_attr "val", type: SimpleTypes::ST_TextSpacingPoint
    end

    # `a:lnSpc`, `a:spcBef` and `a:spcAft`: a spacing value, expressed either
    # as a multiple of line height or as a fixed distance.
    class CT_TextSpacing < Element
      tag "a:lnSpc", "a:spcBef", "a:spcAft"
      # Strictly a one-and-only-one choice; modelling it as two optional
      # children is simpler and the setters keep only one present.
      zero_or_one "a:spcPct", successors: []
      zero_or_one "a:spcPts", successors: []

      # Spacing as a multiple of line height, e.g. 1.75 lines.
      def set_spc_pct(value)
        remove_spcPts
        get_or_add_spcPct.val = value
      end

      # Spacing as a fixed distance.
      def set_spc_pts(value)
        remove_spcPct
        get_or_add_spcPts.val = value
      end
    end

    # `a:pPr`, paragraph properties.
    class CT_TextParagraphProperties < Element
      tag "a:pPr"
      TAG_SEQ = %w[a:lnSpc a:spcBef a:spcAft a:buClrTx a:buClr a:buSzTx a:buSzPct
                   a:buSzPts a:buFontTx a:buFont a:buNone a:buAutoNum a:buChar
                   a:buBlip a:tabLst a:defRPr a:extLst].freeze

      zero_or_one "a:lnSpc", successors: TAG_SEQ[1..]
      zero_or_one "a:spcBef", successors: TAG_SEQ[2..]
      zero_or_one "a:spcAft", successors: TAG_SEQ[3..]
      zero_or_one "a:defRPr", successors: TAG_SEQ[16..]
      optional_attr "lvl", type: SimpleTypes::ST_TextIndentLevelType, default: 0
      optional_attr "algn", type: Enum::PP_PARAGRAPH_ALIGNMENT

      # A Float means a number of lines; a {Pptx::Length} means a fixed
      # distance. nil when no spacing is set.
      def line_spacing
        spacing = lnSpc
        return nil if spacing.nil?

        spacing.spcPts ? spacing.spcPts.val : spacing.spcPct&.val
      end

      def line_spacing=(value)
        remove_lnSpc
        return if value.nil?

        spacing = get_or_add_lnSpc
        value.is_a?(Pptx::Length) ? spacing.set_spc_pts(value) : spacing.set_spc_pct(value)
      end

      def space_before = spcBef&.spcPts&.val

      def space_before=(value)
        remove_spcBef
        get_or_add_spcBef.set_spc_pts(value) unless value.nil?
      end

      def space_after = spcAft&.spcPts&.val

      def space_after=(value)
        remove_spcAft
        get_or_add_spcAft.set_spc_pts(value) unless value.nil?
      end
    end

    # `a:p`, a paragraph.
    class CT_TextParagraph < Element
      tag "a:p"
      zero_or_one "a:pPr", successors: %w[a:r a:br a:fld a:endParaRPr]
      zero_or_more "a:r", successors: %w[a:endParaRPr], as: :r
      zero_or_more "a:br", successors: %w[a:endParaRPr], as: :br
      zero_or_one "a:endParaRPr", successors: []

      CONTENT_TAGS = %w[a:r a:br a:fld].freeze

      # A run is created with its required `a:t` already in place.
      def new_r = build_from_xml(%(<a:r #{Ns.nsdecls("a")}><a:t/></a:r>))

      def add_run(text = nil)
        run = add_r
        run.text = text if text && !text.empty?
        run
      end

      def add_line_break = add_br

      # Append +text+ as runs, turning each "\n" or "\v" into a line break.
      #
      # Breaks go *between* items, so leading text never produces one, and an
      # empty segment contributes no run.
      def append_text(text)
        text.to_s.split(/[\n\v]/, -1).each_with_index do |segment, index|
          add_line_break if index.positive?
          add_run(segment) unless segment.empty?
        end
        self
      end

      # The `a:r`, `a:br` and `a:fld` children, in document order.
      def content_children
        tags = CONTENT_TAGS.map { |t| Ns.qn(t) }
        @node.element_children
             .select { |c| tags.include?(Ns.clark_name_of(c)) }
             .map { |c| Element.wrap(c) }
      end

      def text = content_children.map(&:text).join
    end

    # `p:txBody`, `a:txBody` and `c:txPr`: a body of text.
    class CT_TextBody < Element
      # `c:rich` is a text body under a different name: it is what a chart or
      # axis title holds.
      tag "p:txBody", "a:txBody", "c:txPr", "c:rich"
      one_and_only_one "a:bodyPr"
      one_or_more "a:p", successors: [], as: :p

      class << self
        # A `p:txBody` with one empty paragraph, built in +context+'s document.
        def new_element(context) = context.build_from_xml(txbody_xml)

        # A standalone `p:txBody`, for callers with no element to build from.
        def parse_new = Element.parse(txbody_xml)

        def txbody_xml
          <<~XML
            <p:txBody #{Ns.nsdecls("a", "p")}>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p/>
            </p:txBody>
          XML
        end
      end

      # Remove every paragraph, leaving any other children alone.
      def clear_content
        p_list.each { |paragraph| remove(paragraph) }
        self
      end

      # Restore the minimum the schema requires: at least one paragraph.
      def unclear_content
        add_p if p_list.empty?
        self
      end

      # True when the body holds a single empty paragraph.
      def empty?
        paragraphs = p_list
        raise InvalidXmlError, "a text body must have at least one a:p" if paragraphs.empty?

        paragraphs.size == 1 && paragraphs.first.text.empty?
      end
    end
  end
end
