# frozen_string_literal: true

$LOAD_PATH.unshift File.expand_path("../lib", __dir__)

require "ruby_pptx"
require "open3"
require "support/differential"
require "support/schema"

# A constant assigned inside an `RSpec.describe` block lands on Object, not on
# the example group, so two spec files using the same name silently clobber
# each other. Ruby only warns; here that warning is an error.
module FailOnRedefinedConstant
  def warn(message, category: nil, **)
    raise "spec leaked a constant: #{message}" if message.include?("already initialized constant")

    super
  end
end
Warning.extend(FailOnRedefinedConstant)

# CI runs the suite once per XML backend by setting RUBY_PPTX_XML_BACKEND. If
# the library chose differently, that run would quietly test the other one.
if (requested = ENV.fetch("RUBY_PPTX_XML_BACKEND", nil)) && Pptx.xml_backend.to_s != requested.downcase
  abort "RUBY_PPTX_XML_BACKEND=#{requested} but the library loaded #{Pptx.xml_backend}"
end

RSpec.configure do |config|
  config.before(:suite) { puts "XML backend: #{Pptx.xml_backend}" }
  config.expect_with(:rspec) { |c| c.syntax = :expect }
  config.disable_monkey_patching!
  config.order = :random
  Kernel.srand config.seed
end
