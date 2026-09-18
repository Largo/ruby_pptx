# frozen_string_literal: true

require "pptx/opc/package"
require "pptx/opc/constants"
require "pptx/oxml/theme"
require "pptx/theme"

module Pptx
  module Parts
    # A theme part, `/ppt/theme/themeN.xml`.
    #
    # Every slide master relates to exactly one, and it is where the master's
    # colours and fonts actually live -- the master itself only maps them.
    class ThemePart < Opc::XmlPart
      # A new theme part starting from the base theme shipped with the gem.
      def self.new_theme(partname, package, name: "Office Theme")
        part = new(partname, Opc::CONTENT_TYPE::OFC_THEME, package,
                   Oxml::CT_OfficeStyleSheet.new_default)
        part.theme.name = name
        part
      end

      def theme = @theme ||= Theme.new(element, self)
    end
  end
end
