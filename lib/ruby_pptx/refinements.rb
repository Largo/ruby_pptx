# frozen_string_literal: true

require "ruby_pptx/numeric_lengths"

module Pptx
  # Numeric length sugar, scoped to the file that asks for it.
  #
  #   require "ruby_pptx/refinements"
  #   using Pptx::Lengths
  #
  #   slide.shapes.add_shape(:rectangle, at: [1.inch, 2.inch], size: [3.inch, 1.inch])
  #
  # A refinement is the polite version of `pptx/core_ext`: `Numeric` gains
  # these methods inside this file and nowhere else, so nothing else sharing
  # the process can see them.
  module Lengths
    refine Numeric do
      import_methods NumericLengths
    end
  end
end
