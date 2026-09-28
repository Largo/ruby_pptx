# frozen_string_literal: true

require "ruby_pptx/oxml/element"
require "ruby_pptx/oxml/content_model"
require "ruby_pptx/oxml/simple_types"
require "ruby_pptx/enum/shapes"
require "ruby_pptx/enum/text"

module Pptx
  module Oxml
    # Behaviour shared by the shape element classes -- `p:sp`, `p:pic`,
    # `p:cxnSp`, `p:graphicFrame` and `p:grpSp`.
    #
    # Position and size live in an `a:xfrm` grandchild that may be absent, so
    # the accessors here read through to it and create it only on write.
    module BaseShapeElement
      def x = xfrm_attr(:x)

      def x=(value)
        set_xfrm_attr(:x, value)
      end

      def y = xfrm_attr(:y)

      def y=(value)
        set_xfrm_attr(:y, value)
      end

      def cx = xfrm_attr(:cx)

      def cx=(value)
        set_xfrm_attr(:cx, value)
      end

      def cy = xfrm_attr(:cy)

      def cy=(value)
        set_xfrm_attr(:cy, value)
      end

      def flipH = !xfrm_attr(:flipH).nil? && xfrm_attr(:flipH)

      def flipH=(value)
        set_xfrm_attr(:flipH, value)
      end

      def flipV = !xfrm_attr(:flipV).nil? && xfrm_attr(:flipV)

      def flipV=(value)
        set_xfrm_attr(:flipV, value)
      end

      # Clockwise rotation in degrees; 0.0 when not set.
      def rot = xfrm&.rot || 0.0

      def rot=(value)
        get_or_add_xfrm.rot = value
      end

      # The `a:xfrm` grandchild, or nil. `p:grpSp` overrides this, since its
      # transform hangs off `p:grpSpPr` rather than `p:spPr`.
      def xfrm = spPr.xfrm

      def get_or_add_xfrm = spPr.get_or_add_xfrm

      def shape_id = nvXxPr.cNvPr.id

      def shape_name = nvXxPr.cNvPr.name

      def txBody = find("p:txBody")

      def placeholder? = !ph.nil?

      # The `p:ph` descendant marking this as a placeholder, or nil.
      def ph = xpath("./*[1]/p:nvPr/p:ph").first

      def ph_idx = require_ph.idx
      def ph_orient = require_ph.orient
      def ph_sz = require_ph.sz
      def ph_type = require_ph.type

      # The non-visual properties element, whose tag varies by shape type
      # (`p:nvSpPr`, `p:nvPicPr`, ...). It is always the first child.
      def nvXxPr = xpath("./*[1]").first

      private

      def require_ph
        ph || raise(Error, "not a placeholder shape")
      end

      def xfrm_attr(name)
        xfrm&.public_send(name)
      end

      def set_xfrm_attr(name, value)
        get_or_add_xfrm.public_send("#{name}=", value)
        value
      end
    end

    # `a:off` and `a:chOff`: a position, in slide space or in a group's own
    # child coordinate space.
    class CT_Point2D < Element
      tag "a:off", "a:chOff"
      required_attr "x", type: SimpleTypes::ST_Coordinate
      required_attr "y", type: SimpleTypes::ST_Coordinate
    end

    # `a:ext` and `a:chExt`: a size, in slide space or in a group's child space.
    class CT_PositiveSize2D < Element
      tag "a:ext", "a:chExt"
      required_attr "cx", type: SimpleTypes::ST_PositiveCoordinate
      required_attr "cy", type: SimpleTypes::ST_PositiveCoordinate
    end

    # `a:xfrm`, a 2-D transform.
    #
    # Also covers CT_GroupTransform2D, which shares the tag inside a group
    # shape and adds the child-space `a:chOff` and `a:chExt`.
    class CT_Transform2D < Element
      tag "a:xfrm", "p:xfrm"
      zero_or_one "a:off", successors: %w[a:ext a:chOff a:chExt]
      zero_or_one "a:ext", successors: %w[a:chOff a:chExt]
      zero_or_one "a:chOff", successors: %w[a:chExt]
      zero_or_one "a:chExt", successors: []
      optional_attr "rot", type: SimpleTypes::ST_Angle, default: 0.0
      optional_attr "flipH", type: SimpleTypes::XsdBoolean, default: false
      optional_attr "flipV", type: SimpleTypes::XsdBoolean, default: false

      def x = off&.x

      def x=(value)
        get_or_add_off.x = value
      end

      def y = off&.y

      def y=(value)
        get_or_add_off.y = value
      end

      def cx = ext&.cx

      def cx=(value)
        get_or_add_ext.cx = value
      end

      def cy = ext&.cy

      def cy=(value)
        get_or_add_ext.cy = value
      end

      # Both children carry required attributes, so a newly created one must
      # be given zeroes rather than left empty and invalid.
      def new_off
        build("a:off").tap do |off|
          off.x = 0
          off.y = 0
        end
      end

      def new_ext
        build("a:ext").tap do |ext|
          ext.cx = 0
          ext.cy = 0
        end
      end

      def new_chOff
        build("a:chOff").tap do |off|
          off.x = 0
          off.y = 0
        end
      end

      def new_chExt
        build("a:chExt").tap do |ext|
          ext.cx = 0
          ext.cy = 0
        end
      end
    end

    # `p:cNvPr`, the non-visual drawing properties every shape carries.
    class CT_NonVisualDrawingProps < Element
      tag "p:cNvPr"
      zero_or_one "a:hlinkClick", successors: %w[a:hlinkHover a:extLst]
      zero_or_one "a:hlinkHover", successors: %w[a:extLst]
      required_attr "id", type: SimpleTypes::ST_DrawingElementId
      required_attr "name", type: SimpleTypes::XsdString
    end

    # `p:nvPr`, the application non-visual properties, which is where a
    # placeholder declares itself.
    class CT_ApplicationNonVisualDrawingProps < Element
      tag "p:nvPr"
      zero_or_one "p:ph",
                  successors: %w[a:audioCd a:wavAudioFile a:audioFile a:videoFile
                                 a:quickTimeFile p:custDataLst p:extLst]
    end

    # `p:ph`, marking a shape as a placeholder.
    class CT_Placeholder < Element
      tag "p:ph"
      optional_attr "type", type: Enum::PP_PLACEHOLDER, default: Enum::PP_PLACEHOLDER::OBJECT
      optional_attr "orient", type: SimpleTypes::ST_Direction,
                              default: SimpleTypes::ST_Direction::HORZ
      optional_attr "sz", type: SimpleTypes::ST_PlaceholderSize,
                          default: SimpleTypes::ST_PlaceholderSize::FULL
      optional_attr "idx", type: SimpleTypes::XsdUnsignedInt, default: 0
    end

    # `a:ln`, line (outline) properties.
    # `a:prstDash`, one of the preset dash patterns.
    class CT_PresetLineDashProperties < Element
      tag "a:prstDash"
      optional_attr "val", type: Enum::MSO_LINE_DASH_STYLE
    end

    class CT_LineProperties < Element
      tag "a:ln"
      FILL_SUCCESSORS = %w[a:prstDash a:custDash a:round a:bevel a:miter
                           a:headEnd a:tailEnd a:extLst].freeze

      zero_or_one_choice [choice("a:noFill"), choice("a:solidFill"),
                          choice("a:gradFill"), choice("a:pattFill")],
                         successors: FILL_SUCCESSORS, as: :eg_lineFillProperties
      zero_or_one "a:prstDash", successors: FILL_SUCCESSORS[1..]
      zero_or_one "a:custDash", successors: FILL_SUCCESSORS[2..]
      optional_attr "w", type: SimpleTypes::ST_LineWidth, default: Pptx::Length.emu(0)

      # The fill layer names this differently; both point at the same group.
      def eg_fillProperties = eg_lineFillProperties

      def prstDash_val = prstDash&.val

      def prstDash_val=(value)
        remove_custDash
        get_or_add_prstDash.val = value
      end
    end

    # Shape properties: geometry, fill, line and transform.
    #
    # The same content model appears under three prefixes -- `p:spPr` on a
    # shape, `c:spPr` on a chart element, `a:spPr` in DrawingML -- so one
    # class covers all of them.
    class CT_ShapeProperties < Element
      tag "p:spPr", "c:spPr", "a:spPr"
      TAG_SEQ = %w[a:xfrm a:custGeom a:prstGeom a:noFill a:solidFill a:gradFill
                   a:blipFill a:pattFill a:grpFill a:ln a:effectLst a:effectDag
                   a:scene3d a:sp3d a:extLst].freeze

      zero_or_one "a:xfrm", successors: TAG_SEQ[1..]
      zero_or_one "a:custGeom", successors: TAG_SEQ[2..]
      zero_or_one "a:prstGeom", successors: TAG_SEQ[3..]
      zero_or_one_choice [choice("a:noFill"), choice("a:solidFill"), choice("a:gradFill"),
                          choice("a:blipFill"), choice("a:pattFill"), choice("a:grpFill")],
                         successors: TAG_SEQ[9..], as: :eg_fillProperties
      zero_or_one "a:ln", successors: TAG_SEQ[10..]
      zero_or_one "a:effectLst", successors: TAG_SEQ[11..]

      def x = xfrm&.x
      def y = xfrm&.y
      def cx = xfrm&.cx
      def cy = xfrm&.cy

      # A new gradient starts from PowerPoint's default rather than empty; an
      # `a:gradFill` with no stops draws nothing.
      def new_gradFill = CT_GradientFillProperties.new_grad_fill(self)
    end

    # `p:grpSpPr`, the properties of a group shape or shape tree.
    class CT_GroupShapeProperties < Element
      tag "p:grpSpPr"
      TAG_SEQ = %w[a:xfrm a:noFill a:solidFill a:gradFill a:blipFill a:pattFill
                   a:grpFill a:effectLst a:effectDag a:scene3d a:extLst].freeze

      zero_or_one "a:xfrm", successors: TAG_SEQ[1..]
      zero_or_one "a:effectLst", successors: TAG_SEQ[8..]
    end
  end
end
