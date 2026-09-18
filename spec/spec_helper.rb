# frozen_string_literal: true

$LOAD_PATH.unshift File.expand_path("../lib", __dir__)

require "ruby_pptx"
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

RSpec.configure do |config|
  config.expect_with(:rspec) { |c| c.syntax = :expect }
  config.disable_monkey_patching!
  config.order = :random
  Kernel.srand config.seed
end
