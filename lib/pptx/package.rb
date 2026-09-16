# frozen_string_literal: true

require "pptx/opc/package"
require "pptx/opc/constants"
require "pptx/parts/core_properties"
require "pptx/parts/presentation"
require "pptx/parts/slide"

module Pptx
  # A .pptx package: an OPC package that knows about PowerPoint's parts.
  class Package < Opc::OpcPackage
    # The core-properties part, created if the package has none.
    def core_properties
      @core_properties ||= begin
        part_related_by(Opc::RELATIONSHIP_TYPE::CORE_PROPERTIES)
      rescue NotFoundError
        Parts::CorePropertiesPart.default(self).tap do |part|
          relate_to(part, Opc::RELATIONSHIP_TYPE::CORE_PROPERTIES)
        end
      end
    end

    def presentation_part = main_document_part

    # The next free `/ppt/media/imageN.<ext>` partname, reusing gaps.
    def next_image_partname(ext) = next_media_like_partname("image", ext)

    # The next free `/ppt/media/mediaN.<ext>` partname, reusing gaps.
    def next_media_partname(ext) = next_media_like_partname("media", ext)

    private

    def next_media_like_partname(stem, ext)
      prefix = "/ppt/media/#{stem}"
      indexes = parts.filter_map do |part|
        name = part.partname.to_s
        part.partname.idx if name.start_with?(prefix)
      end.sort

      idx = indexes.each_with_index.find { |used, i| (i + 1) < used }&.last&.succ ||
            indexes.size + 1
      Opc::PackURI.new("#{prefix}#{idx}.#{ext}")
    end
  end
end

# Register the part classes for the content types this library models. Anything
# not listed here loads as a plain Opc::Part, which is enough to carry it
# through a round-trip untouched.
{
  Pptx::Opc::CONTENT_TYPE::PML_PRESENTATION_MAIN => Pptx::Parts::PresentationPart,
  Pptx::Opc::CONTENT_TYPE::PML_PRES_MACRO_MAIN => Pptx::Parts::PresentationPart,
  Pptx::Opc::CONTENT_TYPE::PML_SLIDE => Pptx::Parts::SlidePart,
  Pptx::Opc::CONTENT_TYPE::PML_SLIDE_LAYOUT => Pptx::Parts::SlideLayoutPart,
  Pptx::Opc::CONTENT_TYPE::PML_SLIDE_MASTER => Pptx::Parts::SlideMasterPart,
  Pptx::Opc::CONTENT_TYPE::PML_NOTES_MASTER => Pptx::Parts::NotesMasterPart,
  Pptx::Opc::CONTENT_TYPE::PML_NOTES_SLIDE => Pptx::Parts::NotesSlidePart,
  Pptx::Opc::CONTENT_TYPE::OPC_CORE_PROPERTIES => Pptx::Parts::CorePropertiesPart
}.each { |content_type, part_class| Pptx::Opc::PartFactory.register(content_type, part_class) }
