# frozen_string_literal: true

require "pptx/length"

module Pptx
  # Numeric sugar: `1.inch`, `2.5.cm`, `18.pt`.
  #
  # This module only defines the methods. Two files decide how they reach
  # Numeric, and they are the whole choice on offer:
  #
  #   require "pptx/refinements"   # then `using Pptx::Lengths` -- this file only
  #   require "pptx/core_ext"      # Numeric gains them everywhere, process-wide
  #
  # Prefer the refinement. It is scoped to the file that asks for it, so it
  # cannot surprise another library that happens to share the process.
  module NumericLengths
    def emu         = Pptx::Length.emu(self)
    def inches      = Pptx::Length.inches(self)
    def cm          = Pptx::Length.cm(self)
    def mm          = Pptx::Length.mm(self)
    def pt          = Pptx::Length.pt(self)
    def centipoints = Pptx::Length.centipoints(self)

    # Spelled out rather than aliased: `import_methods`, which the refinement
    # uses to share these, cannot carry an alias across.
    def inch   = inches
    def points = pt
    def point  = pt
  end
end
