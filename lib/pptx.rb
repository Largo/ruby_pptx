# frozen_string_literal: true

require "nokogiri"
require "zip"

require "pptx/version"
require "pptx/errors"
require "pptx/length"
require "pptx/oxml/ns"
require "pptx/oxml/simple_types"
require "pptx/oxml/element"
require "pptx/oxml/content_model"
require "pptx/opc/pack_uri"
require "pptx/opc/constants"
require "pptx/opc/spec"
require "pptx/opc/oxml"
require "pptx/opc/serialized"
require "pptx/opc/package"

require "pptx/element_proxy"
require "pptx/oxml/presentation"
require "pptx/oxml/slide"
require "pptx/oxml/core_properties"
require "pptx/slide"
require "pptx/presentation"
require "pptx/parts/core_properties"
require "pptx/parts/slide"
require "pptx/parts/presentation"
require "pptx/package"

require "pptx/enum/base"
require "pptx/enum/action"
require "pptx/enum/chart"
require "pptx/enum/dml"
require "pptx/enum/lang"
require "pptx/enum/shapes"
require "pptx/enum/text"

# A Ruby object model for PowerPoint (.pptx) files, ported from python-pptx.
module Pptx
end
