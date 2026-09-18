# frozen_string_literal: true

require_relative "lib/ruby_pptx/version"

Gem::Specification.new do |spec|
  spec.name     = "ruby_pptx"
  spec.version  = Pptx::VERSION
  spec.authors  = ["Andi"]
  spec.email    = ["andi@idogawa.com"]

  spec.summary  = "Create, read and update PowerPoint (.pptx) files."
  spec.description = "A Ruby port of python-pptx: a full OOXML PresentationML " \
                     "object model with a Ruby-idiomatic public API."
  spec.homepage = "https://github.com/Largo/ruby_pptx"
  spec.license  = "MIT"
  spec.required_ruby_version = ">= 3.3"

  spec.metadata["homepage_uri"]    = spec.homepage
  spec.metadata["source_code_uri"] = spec.homepage
  spec.metadata["changelog_uri"]   = "#{spec.homepage}/blob/main/CHANGELOG.md"
  spec.metadata["rubygems_mfa_required"] = "true"

  spec.files = Dir["lib/**/*.rb", "lib/ruby_pptx/templates/**/*", "LICENSE", "NOTICE", "README.md"]
  spec.require_paths = ["lib"]

  spec.add_dependency "nokogiri", "~> 1.18"
  spec.add_dependency "rubyzip", "~> 3.0"
end
