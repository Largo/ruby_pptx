# frozen_string_literal: true

require "pptx/oxml/action"
require "pptx/enum/action"

module Pptx
  # What happens when a shape or run is clicked.
  #
  # Most callers want {Pptx::BaseShape#hyperlink}, which is a shortcut to this
  # object's URL. Reach for `click_action` when the click does something other
  # than open a link -- jumping to a slide, for instance.
  class ActionSetting
    # `ppaction://hlinkshowjump?jump=...` covers the relative moves.
    RELATIVE_JUMPS = {
      "firstslide" => :FIRST_SLIDE, "lastslide" => :LAST_SLIDE,
      "lastslideviewed" => :LAST_SLIDE_VIEWED, "nextslide" => :NEXT_SLIDE,
      "previousslide" => :PREVIOUS_SLIDE, "endshow" => :END_SHOW
    }.freeze

    # Every other action is identified by the host of its `ppaction://` URL.
    # A hyperlink has no action attribute at all.
    ACTION_VERBS = {
      nil => :HYPERLINK, "hlinksldjump" => :NAMED_SLIDE, "hlinkpres" => :PLAY,
      "hlinkfile" => :OPEN_FILE, "customshow" => :NAMED_SLIDE_SHOW,
      "ole" => :OLE_VERB, "macro" => :RUN_MACRO, "program" => :RUN_PROGRAM
    }.freeze

    # @param x_pr [Pptx::Oxml::Element] a `p:cNvPr` or an `a:rPr`
    # @param parent the shape or run this action belongs to
    def initialize(x_pr, parent, hover: false)
      @element = x_pr
      @parent = parent
      @hover = hover
    end

    def part = @parent.part

    # @return [Pptx::Enum::PP_ACTION] NONE when nothing happens on click
    def action
      link = hlink
      return Enum::PP_ACTION::NONE if link.nil?

      verb = link.action_verb
      if verb == "hlinkshowjump"
        name = RELATIVE_JUMPS[link.action_fields["jump"]]
        return name ? Enum::PP_ACTION.fetch(name) : Enum::PP_ACTION::NONE
      end

      name = ACTION_VERBS[verb]
      name ? Enum::PP_ACTION.fetch(name) : Enum::PP_ACTION::NONE
    end

    # The relationship target of this click, whatever kind it is.
    #
    # For a hyperlink that is the URL. For a slide jump it is the target
    # slide's partname, which is what python-pptx reports here too -- its
    # docstring says otherwise, but the code returns the ref for any action
    # carrying a relationship. {#url} is the one that means "a hyperlink".
    def address
      link = hlink
      return nil if link.nil?

      r_id = link.rId
      return nil if r_id.nil? || r_id.empty?

      part.target_ref(r_id)
    end

    # Set, change or (with nil) remove the hyperlink.
    def address=(url)
      clear
      return if url.nil? || url.empty?

      get_or_add_hlink.rId = part.relate_to(url, Opc::RELATIONSHIP_TYPE::HYPERLINK, external: true)
    end

    # The URL this click opens, or nil when the click is not a hyperlink.
    #
    # Unlike {#address} this does not report a slide partname for a jump: a
    # thing called a URL should be a URL.
    def url = action == Enum::PP_ACTION::HYPERLINK ? address : nil

    # The slide this click jumps to, or nil.
    def target_slide
      return nil unless action == Enum::PP_ACTION::NAMED_SLIDE

      part.related_part(hlink.rId).slide
    end

    def target_slide=(slide)
      clear
      return if slide.nil?

      link = get_or_add_hlink
      link.action = "#{Oxml::CT_Hyperlink::ACTION_SCHEME}hlinksldjump"
      link.rId = part.relate_to(slide.part, Opc::RELATIONSHIP_TYPE::SLIDE)
    end

    # Remove whatever this click does, and the relationship behind it.
    def clear
      link = hlink
      return self if link.nil?

      r_id = link.rId
      part.drop_rel(r_id) if r_id && !r_id.empty?
      @element.remove(link)
      self
    end

    def inspect = "#<Pptx::ActionSetting #{action.name} #{address.inspect}>"

    private

    def hlink = @element.find(@hover ? "a:hlinkHover" : "a:hlinkClick")

    def get_or_add_hlink
      @hover ? @element.get_or_add_hlinkHover : @element.get_or_add_hlinkClick
    end
  end
end
