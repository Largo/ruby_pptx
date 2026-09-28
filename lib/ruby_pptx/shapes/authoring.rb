# frozen_string_literal: true

module Pptx
  # Adding placeholders to a slide master or a slide layout.
  #
  # A slide's placeholders are cloned from its layout and are not created
  # directly, so this is mixed into the master and layout collections only.
  module PlaceholderAuthoring
    # Add a placeholder.
    #
    #   master.placeholders.add(:title, at: [x, y], size: [cx, cy])
    #   layout.placeholders.add(:body, idx: 1, at: [x, y], size: [cx, cy])
    #
    # A title placeholder carries no `idx` -- there is only ever one, and
    # PowerPoint identifies it by type. Every other kind needs one, and gets
    # the next free value unless told otherwise. The `idx` is what ties a
    # layout placeholder to the master placeholder it inherits from, so it is
    # worth setting deliberately when the two should line up.
    #
    # @param ph_type [Symbol, Pptx::Enum::PP_PLACEHOLDER_TYPE]
    # @param at [Array(Length, Length), nil] position; inherited when omitted
    # @param size [Array(Length, Length), nil] size; inherited when omitted
    # @param idx [Integer, nil]
    # @param orient [Symbol, nil] :vertical for a vertical placeholder
    # @param sz [Symbol, nil] :full, :half or :quarter, the size hint
    #   PowerPoint uses when it rebuilds a layout
    # @return [Shape] the new placeholder
    def add(ph_type, at: nil, size: nil, idx: nil, orient: nil, sz: nil)
      type = Enum::PP_PLACEHOLDER_TYPE.fetch(ph_type)
      idx = resolve_idx(type, idx)
      id = next_shape_id
      element = @sp_tree.add_placeholder(id, next_placeholder_name(type, id, orient),
                                         type, orient_value(orient), sz_value(sz), idx)
      shape = shape_factory(element)
      place(shape, at, size)
      shape
    end

    # Add the five placeholders PowerPoint expects a master to have.
    #
    # Their geometry is the Office default scaled to +slide_size+, so a master
    # added to a 4:3 deck reproduces the familiar layout exactly and a 16:9
    # deck gets the same proportions rather than a 4:3 block in the corner.
    def add_standard_set(slide_size)
      width, height = slide_size
      STANDARD_PLACEHOLDERS.each do |ph_type, idx, sz, (x, y, cx, cy)|
        add(ph_type, idx: idx, sz: sz,
                     at: [scale(x, width), scale(y, height)],
                     size: [scale(cx, width), scale(cy, height)])
      end
      self
    end

    # Fractions of the slide, taken from the Office default master, whose
    # slide is 9144000 x 6858000 EMU.
    STANDARD_PLACEHOLDERS = [
      [:TITLE, nil, nil, [457_200 / 9_144_000.0, 274_638 / 6_858_000.0,
                          8_229_600 / 9_144_000.0, 1_143_000 / 6_858_000.0]],
      [:BODY, 1, nil, [457_200 / 9_144_000.0, 1_600_200 / 6_858_000.0,
                       8_229_600 / 9_144_000.0, 4_525_963 / 6_858_000.0]],
      [:DATE, 2, :half, [457_200 / 9_144_000.0, 6_356_350 / 6_858_000.0,
                         2_133_600 / 9_144_000.0, 365_125 / 6_858_000.0]],
      [:FOOTER, 3, :quarter, [3_124_200 / 9_144_000.0, 6_356_350 / 6_858_000.0,
                              2_895_600 / 9_144_000.0, 365_125 / 6_858_000.0]],
      [:SLIDE_NUMBER, 4, :quarter, [6_553_200 / 9_144_000.0, 6_356_350 / 6_858_000.0,
                                    2_133_600 / 9_144_000.0, 365_125 / 6_858_000.0]]
    ].freeze

    SIZES = {
      full: Oxml::SimpleTypes::ST_PlaceholderSize::FULL,
      half: Oxml::SimpleTypes::ST_PlaceholderSize::HALF,
      quarter: Oxml::SimpleTypes::ST_PlaceholderSize::QUARTER
    }.freeze

    private

    def scale(fraction, extent)
      Length.emu((fraction * extent.emu).round)
    end

    def place(shape, at, size)
      shape.left, shape.top = at if at
      shape.width, shape.height = size if size
    end

    # A title is identified by type; everything else by idx, which must be
    # unique within the shape tree.
    def resolve_idx(type, idx)
      return nil if %i[TITLE CENTER_TITLE].include?(type.name)
      return idx unless idx.nil?

      used = member_elements.filter_map(&:ph_idx)
      (1..).find { |candidate| !used.include?(candidate) }
    end

    def orient_value(orient)
      return nil if orient.nil?

      raise ArgumentError, "orient must be :vertical or nil, got #{orient.inspect}" unless orient == :vertical

      Oxml::SimpleTypes::ST_Direction::VERT
    end

    def sz_value(sz)
      return nil if sz.nil?

      SIZES.fetch(sz) do
        raise ArgumentError, "sz must be :full, :half or :quarter, got #{sz.inspect}"
      end
    end
  end
end
