# frozen_string_literal: true

require "ruby_pptx/numeric_lengths"

# Opt-in numeric sugar: `require "ruby_pptx/core_ext"` to write `1.inch` or `2.5.cm`
# instead of `Pptx.inches(1)`. Kept out of the default require so the gem does
# not monkey-patch Numeric behind your back.
#
# This patches Numeric for the whole process. `require "ruby_pptx/refinements"` and
# `using Pptx::Lengths` does the same thing scoped to one file, and is the
# better choice inside a library.
Numeric.include(Pptx::NumericLengths)
