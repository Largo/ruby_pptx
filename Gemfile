# frozen_string_literal: true

source "https://rubygems.org"
gemspec

group :development, :test do
  gem "rake", "~> 13.0"
  gem "rspec", "~> 3.13"
  # Pinned to a patch series on purpose. `.rubocop.yml` sets NewCops: enable,
  # so a minor upgrade would switch on cops nobody has looked at yet and
  # redden CI on a commit that changed nothing.
  gem "rubocop", "~> 1.91.0"
end
