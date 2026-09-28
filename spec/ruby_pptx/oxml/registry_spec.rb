# frozen_string_literal: true

# Specs that define their own element classes register tags globally. If one
# forgets to restore the registry, dispatch silently changes for every spec
# that runs afterwards, which is close to impossible to debug from the
# resulting failure. This asserts the real classes are still in place.
RSpec.describe Pptx::Oxml::Registry do
  {
    "a:xfrm" => Pptx::Oxml::CT_Transform2D,
    "p:xfrm" => Pptx::Oxml::CT_Transform2D,
    "a:off" => Pptx::Oxml::CT_Point2D,
    "a:ext" => Pptx::Oxml::CT_PositiveSize2D,
    "p:sp" => Pptx::Oxml::CT_Shape,
    "p:spTree" => Pptx::Oxml::CT_GroupShape,
    "p:grpSp" => Pptx::Oxml::CT_GroupShape,
    "p:ph" => Pptx::Oxml::CT_Placeholder,
    "p:spPr" => Pptx::Oxml::CT_ShapeProperties,
    "p:txBody" => Pptx::Oxml::CT_TextBody,
    "a:txBody" => Pptx::Oxml::CT_TextBody,
    "p:sld" => Pptx::Oxml::CT_Slide,
    "p:presentation" => Pptx::Oxml::CT_Presentation,
    "cp:coreProperties" => Pptx::Oxml::CT_CoreProperties,
    "pr:Relationships" => Pptx::Opc::Oxml::CT_Relationships,
    "a:p" => Pptx::Oxml::CT_TextParagraph,
    "a:r" => Pptx::Oxml::CT_RegularTextRun,
    "a:rPr" => Pptx::Oxml::CT_TextCharacterProperties,
    "a:endParaRPr" => Pptx::Oxml::CT_TextCharacterProperties,
    "a:bodyPr" => Pptx::Oxml::CT_TextBodyProperties,
    "a:srgbClr" => Pptx::Oxml::CT_SRgbColor,
    "a:schemeClr" => Pptx::Oxml::CT_SchemeColor,
    "a:solidFill" => Pptx::Oxml::CT_SolidColorFillProperties,
    "a:ln" => Pptx::Oxml::CT_LineProperties,
    "a:tbl" => Pptx::Oxml::CT_Table,
    "a:tr" => Pptx::Oxml::CT_TableRow,
    "a:tc" => Pptx::Oxml::CT_TableCell,
    "a:gridCol" => Pptx::Oxml::CT_TableCol,
    "p:pic" => Pptx::Oxml::CT_Picture,
    "p:graphicFrame" => Pptx::Oxml::CT_GraphicalObjectFrame,
    "p:nvPicPr" => Pptx::Oxml::CT_ShapeNonVisualCommon,
    "c:chartSpace" => Pptx::Oxml::CT_ChartSpace,
    "c:externalData" => Pptx::Oxml::CT_ExternalData,
    "p14:sectionLst" => Pptx::Oxml::CT_SectionList,
    "p14:section" => Pptx::Oxml::CT_Section,
    "p:extLst" => Pptx::Oxml::CT_PresentationExtensionList,
    "a:hlinkClick" => Pptx::Oxml::CT_Hyperlink,
    "a:hlinkHover" => Pptx::Oxml::CT_Hyperlink,
    "a:blip" => Pptx::Oxml::CT_Blip,
    "c:chart" => Pptx::Oxml::CT_Chart,
    "c:legend" => Pptx::Oxml::CT_Legend,
    "c:catAx" => Pptx::Oxml::CT_CatAx,
    "c:valAx" => Pptx::Oxml::CT_ValAx,
    "c:barChart" => Pptx::Oxml::CT_Plot,
    "p:cxnSp" => Pptx::Oxml::CT_Connector,
    "p:nvCxnSpPr" => Pptx::Oxml::CT_ConnectorNonVisual,
    "a:stCxn" => Pptx::Oxml::CT_Connection,
    "a:chOff" => Pptx::Oxml::CT_Point2D,
    "a:chExt" => Pptx::Oxml::CT_PositiveSize2D,
    "a:path" => Pptx::Oxml::CT_Path2D,
    "a:moveTo" => Pptx::Oxml::CT_Path2DPoint,
    "a:lnTo" => Pptx::Oxml::CT_Path2DPoint,
    "a:pt" => Pptx::Oxml::CT_AdjPoint2D,
    "p:timing" => Pptx::Oxml::CT_SlideTiming,
    "p:tnLst" => Pptx::Oxml::CT_TimeNodeList,
    "c:ser" => Pptx::Oxml::CT_Series,
    "c:spPr" => Pptx::Oxml::CT_ShapeProperties,
    "c:tx" => Pptx::Oxml::CT_Tx
  }.each do |nsptag, expected_class|
    it "dispatches #{nsptag} to #{expected_class}" do
      expect(described_class.class_for(Pptx::Oxml::Ns.qn(nsptag))).to eq(expected_class)
    end
  end

  # A second class for a tag used to replace the first silently, so the
  # winner depended on load order. c:tx and a:ext both had two classes.
  describe "registering a tag twice" do
    around do |example|
      saved = described_class.registered
      example.run
    ensure
      described_class.reset!(saved)
    end

    it "refuses a second class for the same tag" do
      expect { described_class.register(Pptx::Oxml::Ns.qn("c:tx"), Class.new) }
        .to raise_error(ArgumentError, /already registered to Pptx::Oxml::CT_Tx/)
    end

    it "lets a class register its own tag again" do
      expect { described_class.register(Pptx::Oxml::Ns.qn("c:tx"), Pptx::Oxml::CT_Tx) }
        .not_to raise_error
    end
  end
end
