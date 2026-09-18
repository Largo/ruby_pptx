# frozen_string_literal: true

require "ruby_pptx/element_proxy"
require "ruby_pptx/shapes/base"
require "ruby_pptx/shapes/shape"
require "ruby_pptx/shapes/placeholder"
require "ruby_pptx/autoshape_spec"

module Pptx
  # Builds the right shape proxy for a shape element.
  module ShapeFactory
    module_function

    # The default mapping, used for shapes on a slide master or layout and as
    # the base for the slide-specific variants.
    def build(shape_element, parent)
      case shape_element.nsptag
      when "p:pic"
        shape_element.movie? ? Movie.new(shape_element, parent) : Picture.new(shape_element, parent)
      when "p:cxnSp" then Connector.new(shape_element, parent)
      when "p:grpSp" then GroupShape.new(shape_element, parent)
      when "p:sp" then Shape.new(shape_element, parent)
      when "p:graphicFrame" then GraphicFrame.new(shape_element, parent)
      else BaseShape.new(shape_element, parent)
      end
    end

    # On a slide, a `p:sp` that is a placeholder gets the placeholder proxy.
    def build_for_slide(shape_element, parent)
      return SlidePlaceholder.new(shape_element, parent) if placeholder_sp?(shape_element)

      build(shape_element, parent)
    end

    def build_for_layout(shape_element, parent)
      return LayoutPlaceholder.new(shape_element, parent) if placeholder_sp?(shape_element)

      build(shape_element, parent)
    end

    def build_for_master(shape_element, parent)
      return MasterPlaceholder.new(shape_element, parent) if placeholder_sp?(shape_element)

      build(shape_element, parent)
    end

    def placeholder_sp?(shape_element)
      shape_element.nsptag == "p:sp" && shape_element.placeholder?
    end
  end

  # The shapes of a slide, layout or master, in z-order: first is backmost.
  class BaseShapes < ParentedElementProxy
    include Enumerable

    def initialize(sp_tree, parent)
      super
      @sp_tree = sp_tree
    end

    def each
      return enum_for(:each) { size } unless block_given?

      member_elements.each { |element| yield shape_factory(element) }
      self
    end

    # A group counts as one shape, whatever it contains.
    def size = member_elements.size
    alias length size

    def [](index, length = nil)
      slice_members(member_elements, index, length) { |element| shape_factory(element) }
    end

    def fetch(index)
      self[index] || raise(IndexError, "shape index #{index} out of range")
    end

    # Add a placeholder to this collection modelled on +placeholder+.
    def clone_placeholder(placeholder)
      sp = placeholder.element
      id = next_shape_id
      name = next_placeholder_name(sp.ph_type, id, sp.ph_orient)
      @sp_tree.add_placeholder(id, name, sp.ph_type, sp.ph_orient, sp.ph_sz, sp.ph_idx)
    end

    # The base name PowerPoint gives a placeholder of +ph_type+.
    #
    # A notes slide names its body placeholder differently, so subclasses can
    # override this.
    def placeholder_basename(ph_type)
      PLACEHOLDER_BASENAMES.fetch(ph_type.name) do
        raise ArgumentError, "no placeholder base name for #{ph_type}"
      end
    end

    PLACEHOLDER_BASENAMES = {
      BITMAP: "ClipArt Placeholder",
      BODY: "Text Placeholder",
      CENTER_TITLE: "Title",
      CHART: "Chart Placeholder",
      DATE: "Date Placeholder",
      FOOTER: "Footer Placeholder",
      HEADER: "Header Placeholder",
      MEDIA_CLIP: "Media Placeholder",
      OBJECT: "Content Placeholder",
      ORG_CHART: "SmartArt Placeholder",
      PICTURE: "Picture Placeholder",
      SLIDE_NUMBER: "Slide Number Placeholder",
      SUBTITLE: "Subtitle",
      TABLE: "Table Placeholder",
      TITLE: "Title"
    }.freeze

    private

    # Decide which file the blip points at and which, if any, hangs off its
    # SVG extension.
    #
    # @return [Array(Object, Object)] the raster to embed, and the SVG or nil
    def vector_or_raster(image_file, fallback)
      unless Image.from_file(image_file).vector?
        return [image_file, nil] if fallback.nil?

        raise ArgumentError, "fallback: is only meaningful for a vector image"
      end

      if fallback.nil?
        raise ArgumentError,
              "an SVG needs a raster fallback: add_picture(svg, fallback: png). " \
              "This gem cannot rasterize one for you."
      end

      [fallback, image_file]
    end

    # Which shape elements belong to this collection; placeholder collections
    # narrow this.
    def member?(_shape_element) = true

    def member_elements = @sp_tree.shape_elements.select { |e| member?(e) }

    def shape_factory(shape_element) = ShapeFactory.build(shape_element, self)

    # One more than the highest id in use. Note the shape tree's own allocator
    # fills gaps instead; both behaviours are inherited from python-pptx.
    def next_shape_id = @sp_tree.max_shape_id + 1

    # The name PowerPoint would give a new placeholder: the base name for its
    # type followed by id - 1, bumped until it is unique in the tree, and
    # prefixed "Vertical " for a vertical placeholder.
    def next_placeholder_name(ph_type, id, orient)
      basename = placeholder_basename(ph_type)
      basename = "Vertical #{basename}" if orient == Oxml::SimpleTypes::ST_Direction::VERT

      taken = @sp_tree.xpath("//p:cNvPr/@name").map(&:value)
      numpart = id - 1
      numpart += 1 while taken.include?("#{basename} #{numpart}")
      "#{basename} #{numpart}"
    end
  end

  # A shape collection you can add shapes to: a slide's tree, or a group
  # within one.
  #
  # Adding to a group changes the group's extents, so every method here ends
  # by recalculating them. On a slide that is a no-op.
  class BaseGroupShapes < BaseShapes
    # Add an auto shape.
    #
    #   shapes.add_shape(:rounded_rectangle,
    #                    at: [Pptx.inches(1), Pptx.inches(1)],
    #                    size: [Pptx.inches(2), Pptx.inches(1)])
    #
    # @param shape_type [Pptx::Enum::MSO_SHAPE, Symbol, Integer]
    # @param at [Array(Length, Length)] left and top
    # @param size [Array(Length, Length)] width and height
    # @return [Shape]
    def add_shape(shape_type, at:, size:)
      left, top = at
      width, height = size
      member = Enum::MSO_SHAPE.fetch(shape_type)
      id = next_shape_id
      name = "#{AutoShapeSpec.basename(member)} #{id - 1}"
      sp = @sp_tree.add_autoshape(id, name, Enum::MSO_SHAPE.to_xml(member),
                                  left, top, width, height)
      recalculate_extents
      shape_factory(sp)
    end

    # Add a picture showing the image in +image_file+, which may be a path or
    # an IO stream.
    #
    # Supplying neither +width+ nor +height+ uses the image's native size;
    # supplying one scales the other to preserve the aspect ratio; supplying
    # both stretches the image to fit.
    #
    # An SVG additionally needs +fallback+: a raster image PowerPoint shows to
    # consumers that cannot draw the vector. This gem has no rasterizer, so
    # the fallback has to be supplied rather than generated. The picture is
    # sized from the fallback, since an SVG carries no pixel size of its own.
    #
    #   shapes.add_picture("logo.svg", at: [x, y], fallback: "logo.png")
    #
    # @param at [Array(Length, Length)] left and top
    # @param fallback [String, IO, nil] raster stand-in, required for an SVG
    # @return [Picture]
    def add_picture(image_file, at:, width: nil, height: nil, fallback: nil)
      left, top = at
      raster_file, svg_file = vector_or_raster(image_file, fallback)

      image_part, r_id = part.get_or_add_image_part(raster_file)
      scaled_width, scaled_height = image_part.scale(width, height)
      id = next_shape_id
      pic = @sp_tree.add_pic(id, "Picture #{id - 1}", image_part.desc, r_id,
                             left, top, scaled_width, scaled_height)

      if svg_file
        _svg_part, svg_r_id = part.get_or_add_image_part(svg_file)
        pic.blip.add_svg_blip(svg_r_id)
      end

      recalculate_extents
      shape_factory(pic)
    end

    # Add a movie showing the video in +movie_file+.
    #
    # The size must be given: unlike a picture, a video is not interrogated
    # for its dimensions. Nor is it interrogated for its type, so say what it
    # is with +content_type+; PowerPoint plays "video/unknown" anyway, which
    # is why that is the default.
    #
    # +poster_frame+ is the still shown before the video plays. Without one,
    # the loudspeaker image PowerPoint uses is supplied.
    #
    # @return [Movie]
    def add_movie(movie_file, at:, size:, poster_frame: nil,
                  content_type: Video::UNKNOWN_CONTENT_TYPE)
      left, top = at
      width, height = size
      video = Video.from_file(movie_file, content_type)
      media_r_id, video_r_id = part.get_or_add_video_media_part(video)
      _poster_part, poster_r_id = part.get_or_add_image_part(poster_frame || default_poster_frame)

      id = next_shape_id
      pic = @sp_tree.add_video_pic(id, video.filename, video_r_id, media_r_id, poster_r_id,
                                   left, top, width, height)
      register_video_timing(pic)
      recalculate_extents
      shape_factory(pic)
    end

    # Add a chart of +chart_type+ depicting +chart_data+.
    #
    # @return [GraphicFrame] use its `#chart` to reach the chart itself
    def add_chart(chart_type, chart_data, at:, size:)
      left, top = at
      width, height = size
      r_id = part.add_chart_part(chart_type, chart_data)
      id = next_shape_id
      frame = @sp_tree.add_graphic_frame_chart(id, "Chart #{id - 1}", r_id,
                                               left, top, width, height)
      recalculate_extents
      shape_factory(frame)
    end

    # Add a chart drawing several plots over the same categories.
    #
    #   shapes.add_combo_chart(data, at: [x, y], size: [w, h]) do |combo|
    #     combo.plot :column_clustered, series: "Revenue"
    #     combo.plot :line, series: "Margin", secondary_axis: true
    #   end
    #
    # All the series live in one ChartData, since they share a worksheet; each
    # plot names the ones it draws. PowerPoint can do this and python-pptx
    # cannot, so there is no reference implementation to compare against.
    #
    # @return [GraphicFrame] use its `#chart` to reach the chart itself
    def add_combo_chart(chart_data, at:, size:)
      builder = ComboChartBuilder.new(chart_data)
      yield builder if block_given?

      left, top = at
      width, height = size
      r_id = part.add_combo_chart_part(builder)
      id = next_shape_id
      frame = @sp_tree.add_graphic_frame_chart(id, "Chart #{id - 1}", r_id,
                                               left, top, width, height)
      recalculate_extents
      shape_factory(frame)
    end

    # Add a table of +rows+ by +cols+.
    #
    # @return [GraphicFrame] use its `#table` to reach the table itself
    def add_table(rows, cols, at:, size:)
      left, top = at
      width, height = size
      id = next_shape_id
      frame = @sp_tree.add_graphic_frame_table(id, "Table #{id - 1}", rows, cols,
                                               left, top, width, height)
      recalculate_extents
      shape_factory(frame)
    end

    # Add an empty text box.
    #
    # @return [Shape]
    def add_textbox(at:, size:)
      left, top = at
      width, height = size
      id = next_shape_id
      sp = @sp_tree.add_textbox(id, "TextBox #{id - 1}", left, top, width, height)
      recalculate_extents
      shape_factory(sp)
    end

    # Add a connector between two points.
    #
    # A connector is stored as a bounding box with flip flags rather than as
    # two points, so the end points are converted here.
    #
    # @return [Connector]
    def add_connector(connector_type, begin_at:, end_at:)
      member = Enum::MSO_CONNECTOR_TYPE.fetch(connector_type)
      begin_x, begin_y = begin_at.map { |value| Length.coerce(value).emu }
      end_x, end_y = end_at.map { |value| Length.coerce(value).emu }

      id = next_shape_id
      cxn_sp = @sp_tree.add_cxnSp(
        id, "Connector #{id - 1}", Enum::MSO_CONNECTOR_TYPE.to_xml(member),
        [begin_x, end_x].min, [begin_y, end_y].min,
        (end_x - begin_x).abs, (end_y - begin_y).abs,
        begin_x > end_x, begin_y > end_y
      )
      recalculate_extents
      shape_factory(cxn_sp)
    end

    # Start building a freeform shape.
    #
    # +scale+ says how many EMU one local coordinate unit is worth, so a shape
    # can be described in convenient numbers; pass a pair for different
    # horizontal and vertical scales.
    #
    # @return [FreeformBuilder]
    def build_freeform(start_x: 0, start_y: 0, scale: 1.0)
      FreeformBuilder.new_builder(self, start_x, start_y, scale)
    end

    # Build a freeform shape and add it, in one call.
    #
    #   shapes.add_freeform(at: [x, y], scale: Pptx.inches(1).emu / 100.0) do |f|
    #     f.line_to(100, 0)
    #     f.line_to(50, 100)
    #   end
    #
    # @return [Shape]
    def add_freeform(at: [0, 0], start_x: 0, start_y: 0, scale: 1.0, close: true)
      builder = build_freeform(start_x: start_x, start_y: start_y, scale: scale)
      yield builder if block_given?
      builder.close if close
      builder.convert_to_shape(origin_at: at)
    end

    # The loudspeaker still PowerPoint shows for a video with no poster frame.
    def default_poster_frame
      StringIO.new(File.binread(
                     File.expand_path("../templates/media-speaker.png", __dir__)
                   ))
    end

    # Play controls appear only for a movie listed in the slide's timing tree.
    def register_video_timing(pic)
      slide_element = @sp_tree.xpath("/p:sld").first
      return if slide_element.nil?

      slide_element.get_or_add_child_time_node_list.add_video(pic.shape_id)
    end

    # @api private
    # Used by FreeformBuilder, which needs to add the element and then draw
    # into it.
    def add_freeform_element(x, y, width, height)
      @sp_tree.add_freeform_sp(x, y, width, height)
    end

    # @api private
    def build_shape(shape_element)
      recalculate_extents
      shape_factory(shape_element)
    end

    # Add a group, optionally moving +shapes+ into it.
    #
    # The group has no position or size of its own: both follow from what it
    # contains, and are recomputed whenever its contents change.
    #
    # @return [GroupShape]
    def add_group_shape(shapes = [])
      grp_sp = @sp_tree.add_grpSp
      shapes.each { |shape| grp_sp.insert_element_before(shape.element, "p:extLst") }
      grp_sp.recalculate_extents unless shapes.empty?
      recalculate_extents
      shape_factory(grp_sp)
    end

    private

    # A group resizes itself around its contents; a slide does not move.
    def recalculate_extents = nil

    def shape_factory(shape_element) = ShapeFactory.build_for_slide(shape_element, self)
  end

  # The shapes on a slide.
  class SlideShapes < BaseGroupShapes
    # Copy the layout's placeholders onto this slide, preserving z-order.
    #
    # Latent placeholders -- date, footer and slide number -- are not cloned:
    # PowerPoint shows them from the layout without materialising them on the
    # slide.
    def clone_layout_placeholders(slide_layout)
      slide_layout.cloneable_placeholders.each { |ph| clone_placeholder(ph) }
      self
    end

    # The title placeholder, which is the one with idx 0.
    #
    # @return [SlidePlaceholder, nil]
    def title
      element = @sp_tree.placeholder_elements.find { |e| e.ph_idx.zero? }
      element && shape_factory(element)
    end

    def placeholders = parent.placeholders

    private

    def shape_factory(shape_element) = ShapeFactory.build_for_slide(shape_element, self)
  end

  # The shapes inside a `p:grpSp`.
  class GroupShapes < BaseGroupShapes
    private

    # Adding to a group changes where the group sits and how big it is.
    def recalculate_extents
      @sp_tree.recalculate_extents
      nil
    end
  end

  # The shapes on a slide layout.
  class LayoutShapes < BaseShapes
    private

    def shape_factory(shape_element) = ShapeFactory.build_for_layout(shape_element, self)
  end

  # The shapes on a slide master.
  class MasterShapes < BaseShapes
    private

    def shape_factory(shape_element) = ShapeFactory.build_for_master(shape_element, self)
  end

  # The placeholders of a slide layout, in `idx` order.
  class LayoutPlaceholders < LayoutShapes
    include PlaceholderAuthoring

    # @return [LayoutPlaceholder, nil]
    def by_idx(idx) = find { |ph| ph.element.ph_idx == idx }

    private

    def member?(shape_element) = shape_element.placeholder?

    def member_elements = super.sort_by(&:ph_idx)
  end

  # The placeholders of a slide master, in `idx` order.
  class MasterPlaceholders < MasterShapes
    include PlaceholderAuthoring

    # @return [MasterPlaceholder, nil]
    def by_type(ph_type) = find { |ph| ph.element.ph_type == ph_type }

    private

    def member?(shape_element) = shape_element.placeholder?

    def member_elements = super.sort_by(&:ph_idx)
  end

  # The placeholders of a slide.
  #
  # Ordered by `idx` and looked up by it: `placeholders[1]` is the placeholder
  # whose idx is 1, not the second one in the collection.
  class SlidePlaceholders < ParentedElementProxy
    include Enumerable

    def initialize(sp_tree, parent)
      super
      @sp_tree = sp_tree
    end

    def each
      return enum_for(:each) { size } unless block_given?

      placeholder_elements.each { |e| yield ShapeFactory.build_for_slide(e, self) }
      self
    end

    def size = @sp_tree.placeholder_elements.count
    alias length size

    # The placeholder whose `idx` is +idx+, or nil.
    def [](idx)
      element = @sp_tree.placeholder_elements.find { |e| e.ph_idx == idx }
      element && ShapeFactory.build_for_slide(element, self)
    end

    def fetch(idx)
      self[idx] || raise(NotFoundError, "no placeholder on this slide with idx #{idx}")
    end

    private

    def placeholder_elements = @sp_tree.placeholder_elements.sort_by(&:ph_idx)
  end
end
