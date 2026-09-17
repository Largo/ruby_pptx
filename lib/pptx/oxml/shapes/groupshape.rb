# frozen_string_literal: true

require "pptx/oxml/shapes/shared"
require "pptx/oxml/shapes/autoshape"
require "pptx/oxml/shapes/other"

module Pptx
  module Oxml
    # `p:nvGrpSpPr`, the non-visual properties of a group shape.
    class CT_GroupShapeNonVisual < Element
      tag "p:nvGrpSpPr"
      one_and_only_one "p:cNvPr"
    end

    # `p:spTree` and `p:grpSp`: a shape tree or a group of shapes.
    class CT_GroupShape < Element
      include BaseShapeElement
      tag "p:spTree", "p:grpSp"

      one_and_only_one "p:nvGrpSpPr"
      one_and_only_one "p:grpSpPr"

      # The child tags that count as shapes. Anything else in a shape tree
      # (extension lists and so on) is not part of the shape sequence.
      SHAPE_TAGS = %w[p:sp p:grpSp p:graphicFrame p:cxnSp p:pic p:contentPart].freeze

      # A group's transform hangs off `p:grpSpPr`, not `p:spPr`.
      def xfrm = grpSpPr.xfrm

      def get_or_add_xfrm = grpSpPr.get_or_add_xfrm

      def chOff = get_or_add_xfrm.get_or_add_chOff

      def chExt = get_or_add_xfrm.get_or_add_chExt

      # Each shape child, in document order.
      def shape_elements
        return enum_for(:shape_elements) unless block_given?

        tags = SHAPE_TAGS.map { |t| Ns.qn(t) }
        @node.element_children.each do |child|
          yield Element.wrap(child) if tags.include?(Ns.clark_name_of(child))
        end
      end

      # Each placeholder shape child, in document order.
      def placeholder_elements
        return enum_for(:placeholder_elements) unless block_given?

        shape_elements { |shape| yield shape if shape.placeholder? }
      end

      def add_placeholder(id, name, ph_type, orient, sz, idx)
        sp = adopt_xml(CT_Shape.new_placeholder_sp(id, name, ph_type, orient, sz, idx))
        insert_element_before(sp, "p:extLst")
        sp
      end

      def add_autoshape(id, name, prst, x, y, cx, cy)
        sp = adopt_xml(CT_Shape.new_autoshape_sp(id, name, prst, x, y, cx, cy))
        insert_element_before(sp, "p:extLst")
        sp
      end

      def add_pic(id, name, desc, r_id, x, y, cx, cy)
        pic = adopt_xml(CT_Picture.new_pic(id, name, desc, r_id, x, y, cx, cy))
        insert_element_before(pic, "p:extLst")
        pic
      end

      def add_graphic_frame_chart(id, name, r_id, x, y, cx, cy)
        frame = adopt_xml(
          CT_GraphicalObjectFrame.new_chart_graphic_frame(id, name, r_id, x, y, cx, cy)
        )
        insert_element_before(frame, "p:extLst")
        frame
      end

      def add_graphic_frame_table(id, name, rows, cols, x, y, cx, cy)
        frame = adopt_xml(
          CT_GraphicalObjectFrame.new_table_graphic_frame(id, name, rows, cols, x, y, cx, cy)
        )
        insert_element_before(frame, "p:extLst")
        frame
      end

      def add_textbox(id, name, x, y, cx, cy)
        sp = adopt_xml(CT_Shape.new_textbox_sp(id, name, x, y, cx, cy))
        insert_element_before(sp, "p:extLst")
        sp
      end

      # The largest @id anywhere in the document.
      #
      # Ids have document scope and are not used only by shapes, so every @id
      # counts. The minimum in practice is 1, because the `p:spTree` element
      # itself is always id="1".
      def max_shape_id
        used_ids.max || 0
      end

      # The lowest unused positive @id, filling gaps.
      #
      # Note this differs from the shape collection's own allocator, which
      # takes max+1; both behaviours are preserved from python-pptx.
      def next_shape_id
        used = used_ids
        (1..(used.size + 1)).find { |n| !used.include?(n) }
      end

      private

      def used_ids
        xpath("//@id").filter_map do |attr|
          value = attr.value
          Integer(value, 10) if /\A\d+\z/.match?(value)
        end
      end

      # A shape built from an XML literal is parsed into its own document, so
      # it must be rebuilt inside this tree's document before insertion.
      def adopt_xml(element)
        return element if element.document.equal?(document)

        build_from_xml(element.node.to_xml)
      end
    end
  end
end
