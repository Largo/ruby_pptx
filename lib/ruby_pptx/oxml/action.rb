# frozen_string_literal: true

require "ruby_pptx/oxml/element"
require "ruby_pptx/oxml/content_model"
require "ruby_pptx/oxml/simple_types"

module Pptx
  module Oxml
    # `a:hlinkClick` and `a:hlinkHover`: what happens when a shape or run is
    # clicked, or the pointer rests on it.
    #
    # An ordinary web link carries only `r:id`. Everything else PowerPoint can
    # do on click -- jump to a slide, run a macro, open a file -- is encoded in
    # an `action` attribute holding a `ppaction://` URL.
    class CT_Hyperlink < Element
      tag "a:hlinkClick", "a:hlinkHover"
      optional_attr "r:id", type: SimpleTypes::XsdString, as: :rId
      optional_attr "action", type: SimpleTypes::XsdString

      ACTION_SCHEME = "ppaction://"

      # The host part of the action URL, e.g. "hlinksldjump" in
      # "ppaction://hlinksldjump". nil when this is a plain hyperlink.
      def action_verb
        url = action
        return nil if url.nil?

        url.split("?").first.delete_prefix(ACTION_SCHEME)
      end

      # The query of the action URL as a hash, e.g. {"jump" => "nextslide"}.
      def action_fields
        url = action
        return {} if url.nil?

        _host, query = url.split("?", 2)
        return {} if query.nil? || query.empty?

        query.split("&").to_h { |pair| pair.split("=", 2) }
      end
    end
  end
end
