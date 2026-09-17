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

    # The image part holding +image_file+, created if the package has no part
    # with the same content.
    #
    # Images are matched by SHA-1 of their bytes, so the same picture used on
    # several slides is stored once.
    def get_or_add_image_part(image_file)
      image = Image.from_file(image_file)
      find_image_part(image.sha1) || Parts::ImagePart.new_image(self, image)
    end

    # The next free `/ppt/media/imageN.<ext>` partname, reusing gaps.
    def next_image_partname(ext) = next_media_like_partname("image", ext)

    # The next free `/ppt/media/mediaN.<ext>` partname, reusing gaps.
    def next_media_partname(ext) = next_media_like_partname("media", ext)

    private

    def find_image_part(sha1)
      image_parts.find { |part| part.respond_to?(:sha1) && part.sha1 == sha1 }
    end

    # Parts reached by an image relationship, each once.
    #
    # Scoped by relationship rather than by class on purpose: the package
    # thumbnail is also an image part, but it is related as a thumbnail, and
    # matching a new picture against it would relate the picture by the wrong
    # reltype.
    def image_parts
      seen = {}.compare_by_identity
      each_rel.filter_map do |rel|
        next if rel.external? || rel.reltype != Opc::RELATIONSHIP_TYPE::IMAGE

        part = rel.target_part
        next if seen.key?(part)

        seen[part] = true
        part
      end
    end

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
  Pptx::Opc::CONTENT_TYPE::OPC_CORE_PROPERTIES => Pptx::Parts::CorePropertiesPart,
  Pptx::Opc::CONTENT_TYPE::PNG => Pptx::Parts::ImagePart,
  Pptx::Opc::CONTENT_TYPE::JPEG => Pptx::Parts::ImagePart,
  Pptx::Opc::CONTENT_TYPE::GIF => Pptx::Parts::ImagePart,
  Pptx::Opc::CONTENT_TYPE::BMP => Pptx::Parts::ImagePart,
  Pptx::Opc::CONTENT_TYPE::TIFF => Pptx::Parts::ImagePart,
  Pptx::Opc::CONTENT_TYPE::MS_PHOTO => Pptx::Parts::ImagePart,
  Pptx::Opc::CONTENT_TYPE::X_EMF => Pptx::Parts::ImagePart,
  Pptx::Opc::CONTENT_TYPE::X_WMF => Pptx::Parts::ImagePart
}.each { |content_type, part_class| Pptx::Opc::PartFactory.register(content_type, part_class) }
