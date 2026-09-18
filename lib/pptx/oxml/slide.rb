# frozen_string_literal: true

require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/oxml/simple_types"
require "pptx/oxml/ns"

module Pptx
  module Oxml
    # Reads one of the XML templates shipped with the gem, e.g. "notes".
    def self.parse_from_template(name)
      path = File.expand_path("../templates/#{name}.xml", __dir__)
      Element.parse(File.read(path))
    end

    # Behaviour common to the six slide-like root elements.
    module BaseSlideElement
      # The `p:cSld/p:spTree` grandchild every slide type has.
      def spTree = cSld.spTree
    end

    # `p:cSld`, the common slide data holding the shape tree.
    class CT_CommonSlideData < Element
      tag "p:cSld"
      zero_or_one "p:bg", successors: %w[p:spTree p:custDataLst p:controls p:extLst]
      one_and_only_one "p:spTree"
      optional_attr "name", type: SimpleTypes::XsdString, default: ""

      # The `p:bg/p:bgPr` grandchild, establishing a no-fill background first
      # if there is not already an explicit one.
      #
      # This is destructive by nature: a background given by a style reference
      # (`p:bgRef`) is replaced, and an inherited background stops being
      # inherited. python-pptx behaves the same way.
      def get_or_add_bgPr
        background = bg
        background = change_to_no_fill_bg if background.nil? || background.bgPr.nil?
        background.bgPr
      end

      private

      def change_to_no_fill_bg
        remove_bg
        get_or_add_bg.tap(&:add_no_fill_bgPr)
      end
    end

    # `p:bg`, a slide background.
    class CT_Background < Element
      tag "p:bg"
      # Strictly a choice rather than a sequence, but treating it as two
      # optional children is simpler and behaves the same here.
      zero_or_one "p:bgPr", successors: []
      zero_or_one "p:bgRef", successors: []

      def add_no_fill_bgPr
        bgPr = build_from_xml(
          %(<p:bgPr #{Ns.nsdecls("a", "p")}><a:noFill/><a:effectLst/></p:bgPr>)
        )
        insert_bgPr(bgPr)
        bgPr
      end
    end

    # `p:bgPr`, explicit background fill properties.
    class CT_BackgroundProperties < Element
      tag "p:bgPr"
      FILL_SUCCESSORS = %w[a:effectLst a:effectDag a:extLst].freeze

      zero_or_one_choice [choice("a:noFill"), choice("a:solidFill"), choice("a:gradFill"),
                          choice("a:blipFill"), choice("a:pattFill"), choice("a:grpFill")],
                         successors: FILL_SUCCESSORS, as: :eg_fillProperties
    end

    # `p:sld`, the root of a slide part.
    class CT_Slide < Element
      include BaseSlideElement

      tag "p:sld"
      one_and_only_one "p:cSld"
      zero_or_one "p:clrMapOvr", successors: %w[p:transition p:timing p:extLst]
      zero_or_one "p:timing", successors: %w[p:extLst]

      # A new, empty `p:sld` with the minimum structure PowerPoint requires.
      def self.new_element = Element.parse(SLD_XML)

      def bg = cSld.bg

      # The node list a `p:video` entry hangs off, which is what makes
      # PowerPoint show play controls under a movie.
      #
      # The precondition varies -- a slide may have no timing at all, or one
      # already -- so an awkward case is replaced wholesale rather than
      # patched. That can in principle discard existing timing information;
      # python-pptx takes the same view.
      def get_or_add_child_time_node_list
        existing = xpath("./p:timing/p:tnLst/p:par/p:cTn/p:childTnLst").first
        return existing if existing

        remove(get_or_add_timing)
        timing = build_from_xml(TIMING_XML)
        insert_timing(timing)
        timing.xpath("./p:tnLst/p:par/p:cTn/p:childTnLst").first
      end

      TIMING_XML = <<~XML.freeze
        <p:timing #{Ns.nsdecls("p")}>
          <p:tnLst>
            <p:par>
              <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                <p:childTnLst/>
              </p:cTn>
            </p:par>
          </p:tnLst>
        </p:timing>
      XML

      SLD_XML = <<~XML.freeze
        <p:sld #{Ns.nsdecls("a", "p", "r")}>
          <p:cSld>
            <p:spTree>
              <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
              </p:nvGrpSpPr>
              <p:grpSpPr/>
            </p:spTree>
          </p:cSld>
          <p:clrMapOvr>
            <a:masterClrMapping/>
          </p:clrMapOvr>
        </p:sld>
      XML
    end

    # `p:sldLayout`, the root of a slide-layout part.
    class CT_SlideLayout < Element
      include BaseSlideElement

      tag "p:sldLayout"
      one_and_only_one "p:cSld"
      zero_or_one "p:clrMapOvr", successors: %w[p:transition p:timing p:hf p:extLst]
      optional_attr "type", type: SimpleTypes::ST_SlideLayoutType
      optional_attr "preserve", type: SimpleTypes::XsdBoolean

      # A new `p:sldLayout` named +name+, inheriting its colour map from the
      # master. `preserve` is set because a layout with no slides using it is
      # otherwise discarded by PowerPoint when it tidies up.
      def self.new_element(name, layout_type)
        layout = Element.parse(SLD_LAYOUT_XML)
        layout.cSld.name = name
        layout.type = layout_type unless layout_type.nil?
        layout
      end

      SLD_LAYOUT_XML = <<~XML.freeze
        <p:sldLayout #{Ns.nsdecls("a", "p", "r")} preserve="1">
          <p:cSld>
            <p:spTree>
              <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
              </p:nvGrpSpPr>
              <p:grpSpPr/>
            </p:spTree>
          </p:cSld>
          <p:clrMapOvr>
            <a:masterClrMapping/>
          </p:clrMapOvr>
        </p:sldLayout>
      XML
    end

    # `p:sldMaster`, the root of a slide-master part.
    class CT_SlideMaster < Element
      include BaseSlideElement

      tag "p:sldMaster"
      one_and_only_one "p:cSld"
      one_and_only_one "p:clrMap"
      zero_or_one "p:sldLayoutIdLst",
                  successors: %w[p:transition p:timing p:hf p:txStyles p:extLst]
      zero_or_one "p:txStyles", successors: %w[p:extLst]

      # A new `p:sldMaster` carrying the colour map and text styles a master
      # cannot do without, an empty shape tree and no layouts.
      def self.new_element = Oxml.parse_from_template("slideMaster")
    end

    # `p:clrMap`, which binds each of a slide's colour roles to a theme colour.
    class CT_ColorMapping < Element
      tag "p:clrMap"
      ROLES = %w[bg1 tx1 bg2 tx2 accent1 accent2 accent3 accent4 accent5
                 accent6 hlink folHlink].freeze
      ROLES.each { |role| required_attr role, type: SimpleTypes::XsdString }
    end

    # `p:txStyles`, the master's default text styling per outline level.
    class CT_SlideMasterTextStyles < Element
      tag "p:txStyles"
      zero_or_one "p:titleStyle", successors: %w[p:bodyStyle p:otherStyle p:extLst]
      zero_or_one "p:bodyStyle", successors: %w[p:otherStyle p:extLst]
      zero_or_one "p:otherStyle", successors: %w[p:extLst]
    end

    # `p:sldLayoutIdLst`, the layouts inheriting from a slide master.
    class CT_SlideLayoutIdList < Element
      tag "p:sldLayoutIdLst"
      zero_or_more "p:sldLayoutId", as: :sldLayoutId

      def size = sldLayoutId_list.size
    end

    # `p:sldLayoutId`, a reference to one slide layout.
    class CT_SlideLayoutIdListEntry < Element
      tag "p:sldLayoutId"
      required_attr "r:id", type: SimpleTypes::XsdString, as: :rId
      optional_attr "id", type: SimpleTypes::ST_SlideLayoutId
    end

    # `p:notesMaster`, the root of the notes-master part.
    class CT_NotesMaster < Element
      include BaseSlideElement

      tag "p:notesMaster"
      one_and_only_one "p:cSld"

      def self.new_default = Oxml.parse_from_template("notesMaster")
    end

    # `p:notes`, the root of a notes-slide part.
    #
    # The template carries no placeholders; those are cloned from the notes
    # master once the shape layer exists.
    class CT_NotesSlide < Element
      include BaseSlideElement

      tag "p:notes"
      one_and_only_one "p:cSld"

      def self.new_element = Oxml.parse_from_template("notes")
    end

    # `p:timing`, animation and timing information.
    class CT_SlideTiming < Element
      tag "p:timing"
      zero_or_one "p:tnLst", successors: %w[p:bldLst p:extLst]
    end

    # `p:tnLst` and `p:childTnLst`, lists of timing nodes.
    class CT_TimeNodeList < Element
      tag "p:tnLst", "p:childTnLst"

      # Register the movie with shape id +shape_id+ so its play controls
      # appear.
      def add_video(shape_id)
        append(build_from_xml(<<~XML))
          <p:video #{Ns.nsdecls("p")}>
            <p:cMediaNode vol="80000">
              <p:cTn id="#{next_time_node_id}" fill="hold" display="0">
                <p:stCondLst>
                  <p:cond delay="indefinite"/>
                </p:stCondLst>
              </p:cTn>
              <p:tgtEl>
                <p:spTgt spid="#{shape_id}"/>
              </p:tgtEl>
            </p:cMediaNode>
          </p:video>
        XML
        self
      end

      private

      # Timing node ids are scoped to the slide, so every existing one counts.
      def next_time_node_id
        used = xpath("/p:sld/p:timing//p:cTn/@id").map { |attr| attr.value.to_i }
        (used.max || 0) + 1
      end
    end
  end
end
