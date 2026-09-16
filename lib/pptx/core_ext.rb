# frozen_string_literal: true

require "pptx/length"

# Opt-in numeric sugar: `require "pptx/core_ext"` to write `1.inch` or `2.5.cm`
# instead of `Pptx.inches(1)`. Kept out of the default require so the gem does
# not monkey-patch Numeric behind your back.
module Pptx
  module NumericLengths
    def emu         = Pptx::Length.emu(self)
    def inches      = Pptx::Length.inches(self)
    def cm          = Pptx::Length.cm(self)
    def mm          = Pptx::Length.mm(self)
    def pt          = Pptx::Length.pt(self)
    def centipoints = Pptx::Length.centipoints(self)

    alias inch inches
    alias points pt
    alias point pt
  end
end

Numeric.include(Pptx::NumericLengths)
