require 'rspec'

RSpec.configure do |config|
  # enable should syntax
  config.expect_with(:rspec) { |c| c.syntax = :should }
  config.example_status_persistence_file_path = "spec/examples.txt"
end

# implement minimal minitest support for migrated tests
class RSpec::Core::ExampleGroup
  def assert_equal a, b
    a.should == b
  end

  def assert pred
    pred.should be_truthy
  end
end
