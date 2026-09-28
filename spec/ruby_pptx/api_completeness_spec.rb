# frozen_string_literal: true

require "json"
require_relative "../../tools/api_audit"

# The port is complete when every public member of python-pptx's API has a
# counterpart here. This was once claimed without being checked, and was
# wrong by 150 members; now the claim is this spec.
RSpec.describe "coverage of python-pptx's public API" do
  before { require_oracle! }

  let(:audit) do
    dump, err, status = Open3.capture3("python3", File.expand_path("../../tools/api_dump.py", __dir__))
    raise "api dump failed: #{err}" unless status.success?

    ApiAudit.run(JSON.parse(dump))
  end

  # The audit is pinned to the python-pptx the rest of the suite uses.
  it "audits the pinned python-pptx" do
    expect(audit["version"]).to eq("1.0.2")
  end

  it "has a counterpart for every public class" do
    expect(audit["missing_classes"]).to eq([])
  end

  it "has a counterpart for every public member" do
    expect(audit["missing_members"]).to eq({})
  end

  # A class the audit neither checks nor excuses is a class it cannot see,
  # which is how multi-level chart categories once went unnoticed.
  it "accounts for every public class, checked or excused" do
    expect(audit["unaccounted"]).to eq([])
  end
end
