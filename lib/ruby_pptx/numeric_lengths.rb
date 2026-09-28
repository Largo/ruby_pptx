# frozen_string_literal: true

require "ruby_pptx/length"

module Pptx
  # Numeric sugar: `1.inch`, `2.5.cm`, `18.pt`.
  #
  # This module only defines the methods. Two files decide how they reach
  # Numeric, and they are the whole choice on offer:
  #
  #   require "ruby_pptx/refinements"   # then `using Pptx::Lengths` -- this file only
  #   require "ruby_pptx/core_ext"      # Numeric gains them everywhere, process-wide
  #
  # Prefer the refinement. It is scoped to the file that asks for it, so it
  # cannot surprise another library that happens to share the process.
  module NumericLengths
    def emu
      Pptx::Length.emu(self)
    end

    def inches
      Pptx::Length.inches(self)
    end

    def cm
      Pptx::Length.cm(self)
    end

    def mm
      Pptx::Length.mm(self)
    end

    def pt
      Pptx::Length.pt(self)
    end

    def centipoints
      Pptx::Length.centipoints(self)
    end

    # Spelled out rather than aliased: `import_methods`, which the refinement
    # uses to share these, cannot carry an alias across.
    def inch
      inches
    end

    def points
      pt
    end

    def point
      pt
    end
  end
end
