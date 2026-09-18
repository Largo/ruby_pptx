# frozen_string_literal: true

require "ruby_pptx/opc/package"
require "ruby_pptx/opc/constants"
require "ruby_pptx/oxml/chart"
require "ruby_pptx/chart/chart"

module Pptx
  module Parts
    # Another OPC package embedded inside this one -- here, the Excel workbook
    # that backs a chart.
    class EmbeddedXlsxPart < Opc::Part
      PARTNAME_TEMPLATE = "/ppt/embeddings/Microsoft_Excel_Sheet%d.xlsx"

      def self.new_xlsx(blob, package)
        new(package.next_partname(PARTNAME_TEMPLATE),
            Opc::CONTENT_TYPE::SML_SHEET, package, blob)
      end
    end

    # A chart part, `/ppt/charts/chartN.xml`.
    class ChartPart < Opc::XmlPart
      PARTNAME_TEMPLATE = "/ppt/charts/chart%d.xml"

      # A new chart part depicting +chart_data+, together with the embedded
      # workbook that backs it.
      def self.new_chart(chart_type, chart_data, package)
        part = load(package.next_partname(PARTNAME_TEMPLATE), Opc::CONTENT_TYPE::DML_CHART,
                    package, ChartXmlWriter.write(chart_type, chart_data))
        part.workbook.replace_with(chart_data.xlsx_blob)
        part
      end

      # A chart part drawing several plots over the same categories.
      def self.new_combo_chart(builder, package)
        part = load(package.next_partname(PARTNAME_TEMPLATE), Opc::CONTENT_TYPE::DML_CHART,
                    package, ComboChartWriter.new(builder.validate!).xml)
        part.workbook.replace_with(builder.chart_data.xlsx_blob)
        part
      end

      def chart = @chart ||= Chart.new(element, self)

      def workbook = @workbook ||= ChartWorkbook.new(element, self)
    end

    # The embedded workbook behind a chart.
    class ChartWorkbook
      def initialize(chart_space, chart_part)
        @chart_space = chart_space
        @chart_part = chart_part
      end

      # @return [EmbeddedXlsxPart, nil]
      def xlsx_part
        r_id = @chart_space.xlsx_part_rId
        r_id && @chart_part.related_part(r_id)
      end

      # Put +xlsx_blob+ behind this chart, adding the part and the
      # `c:externalData` link if the chart does not have one yet.
      def replace_with(xlsx_blob)
        existing = xlsx_part
        return existing.blob = xlsx_blob if existing

        part = EmbeddedXlsxPart.new_xlsx(xlsx_blob, @chart_part.package)
        r_id = @chart_part.relate_to(part, Opc::RELATIONSHIP_TYPE::PACKAGE)
        external_data = @chart_space.get_or_add_externalData
        external_data.rId = r_id
        external_data.ensure_auto_update
        part
      end
    end
  end
end
