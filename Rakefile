# frozen_string_literal: true

# build / release, used by the publish workflow (rubygems/release-gem runs
# `rake release`).
require "bundler/gem_tasks"
require "rspec/core/rake_task"

RSpec::Core::RakeTask.new(:spec)
task default: :spec
