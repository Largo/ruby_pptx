# frozen_string_literal: true

require "json"
require "open3"
require "digest"
require "tempfile"
require "tmpdir"

# Compares a package this gem produces against the one python-pptx produces for
# the same operation.
#
# Both sides are reduced to a manifest -- the sorted zip entry list plus one
# canonical form per part (C14N for XML, SHA-256 for binaries) -- so that
# incidental differences in attribute order, whitespace, zip ordering and
# compression level do not register as failures, while any semantic difference
# in the XML does.
#
#   include Pptx::Spec::Differential
#
#   it "matches python-pptx" do
#     expect_same_package(
#       python: 'prs = pptx.Presentation(); prs.save(out)',
#       ruby: ->(path) { Pptx::Presentation.new.save(path) }
#     )
#   end
module Pptx
  module Spec
    module Differential
      ORACLE = File.expand_path("../../tools/pptx_oracle.py", __dir__)

      # True when python-pptx is importable, so specs can skip rather than fail
      # on a machine without it.
      def self.oracle_available? = module_importable?("pptx, lxml")

      # True when openpyxl is importable; only the chart round-trip needs it.
      def self.openpyxl_available? = module_importable?("openpyxl")

      def self.module_importable?(imports)
        @importable ||= {}
        return @importable[imports] if @importable.key?(imports)

        _out, _err, status = Open3.capture3("python3", "-c", "import #{imports}")
        @importable[imports] = status.success?
      end

      # Skip a spec that needs python-pptx -- unless REQUIRE_ORACLE is set, in
      # which case its absence is a failure.
      #
      # CI sets it. Without that, a broken Python environment would quietly
      # reduce the suite to its unit tests while still reporting green, and the
      # differential checks are the part worth having.
      def require_oracle!(what = "python-pptx", available: Differential.oracle_available?)
        return if available

        raise "#{what} is not importable and REQUIRE_ORACLE is set" if ENV["REQUIRE_ORACLE"]

        skip "#{what} not importable"
      end

      # Run +script+ through python-pptx. The script is handed `pptx` and `out`
      # (a path it must save to).
      #
      # @return [Hash] manifest, keys "entries" and "parts"
      def python_pptx_manifest(script)
        out, err, status = Open3.capture3("python3", ORACLE, stdin_data: script)
        raise "oracle failed (#{status.exitstatus}): #{err}" if out.empty?

        result = JSON.parse(out)
        raise "python-pptx raised: #{result["error"]}" unless result["ok"]

        result
      end

      SIMPLE_TYPE_ORACLE = File.expand_path("../../tools/simple_type_oracle.py", __dir__)

      # Evaluate a batch of simple-type conversions with python-pptx.
      #
      # @param cases [Array<Hash>] each {type:, op:, value:}
      # @return [Array<Hash>] {"ok" => true, "kind" =>, "value" =>} or
      #   {"ok" => false, "error" => "ValueError"}
      def python_simple_types(cases)
        out, err, status = Open3.capture3(
          "python3", SIMPLE_TYPE_ORACLE, stdin_data: JSON.dump(cases)
        )
        raise "simple-type oracle failed (#{status.exitstatus}): #{err}" if out.empty?

        JSON.parse(out)
      end

      # Reduce a .pptx this gem wrote to the same manifest shape.
      def ruby_pptx_manifest(path)
        entries = []
        parts = {}
        Zip::File.open(path) do |zip|
          zip.each do |entry|
            next unless entry.file?

            entries << entry.name
            parts[entry.name] = canonicalize(entry.name, entry.get_input_stream.read)
          end
        end
        { "entries" => entries.sort, "parts" => parts.sort.to_h }
      end

      def canonicalize(name, data)
        return { "kind" => "binary", "sha" => Digest::SHA256.hexdigest(data) } unless xml?(name)

        doc = Nokogiri::XML(data, &:noblanks)
        if doc.errors.any?
          return { "kind" => "malformed", "error" => doc.errors.first.to_s,
                   "raw" => [data].pack("m0") }
        end

        { "kind" => "xml", "c14n" => doc.canonicalize(Nokogiri::XML::XML_C14N_1_1) }
      end

      def xml?(name) = name.downcase.end_with?(".xml", ".rels")

      # The main assertion. +python:+ is a snippet for the oracle; +ruby:+ is a
      # lambda handed a path to write to.
      def expect_same_package(python:, ruby:, ignore: [])
        skip "python-pptx not importable" unless Differential.oracle_available?

        expected = python_pptx_manifest(python)
        actual = Tempfile.create(["ruby_pptx", ".pptx"]) do |file|
          file.close
          ruby.call(file.path)
          ruby_pptx_manifest(file.path)
        end

        compare_manifests(expected, actual, ignore)
      end

      def compare_manifests(expected, actual, ignore)
        ignore = ignore.map(&:to_s)
        keep = ->(list) { list.reject { |n| ignore.include?(n) } }

        missing = keep.call(expected["entries"]) - actual["entries"]
        extra   = keep.call(actual["entries"]) - expected["entries"]
        raise_entry_mismatch(missing, extra) if missing.any? || extra.any?

        keep.call(expected["entries"]).each do |name|
          next if expected["parts"][name] == actual["parts"][name]

          raise part_diff_message(name, expected["parts"][name], actual["parts"][name])
        end
        true
      end

      def raise_entry_mismatch(missing, extra)
        raise RSpec::Expectations::ExpectationNotMetError, <<~MSG
          package entries differ from python-pptx
            missing from ours: #{missing.inspect}
            not in python-pptx: #{extra.inspect}
        MSG
      end

      def part_diff_message(name, expected, actual)
        RSpec::Expectations::ExpectationNotMetError.new(<<~MSG)
          part #{name} differs from python-pptx

          --- python-pptx
          #{(expected && (expected["c14n"] || expected["sha"])).to_s.lines.first(40).join}

          --- ruby_pptx
          #{(actual && (actual["c14n"] || actual["sha"])).to_s.lines.first(40).join}
        MSG
      end
    end
  end
end

RSpec.configure { |config| config.include Pptx::Spec::Differential }
