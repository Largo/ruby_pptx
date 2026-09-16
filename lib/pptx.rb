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
