require 'rspec'

RSpec.configure do |config|
  # enable should syntax
  config.expect_with(:rspec) { |c| c.syntax = :should }
  config.mock_with(:rspec) { |c| c.syntax = :should }

  # allow for --only-failures
  config.example_status_persistence_file_path = "spec/examples.txt"

  # exclude things marked as performance benchmarks
  config.filter_run_excluding performance: true, display_ui: true, extracted: true

  # to get coverage output in ./coverage/index.html say
  #
  #   rspec -t cov
  #
  # otherwise spec run ignores coverage
  if config.filter[:cov]
    config.filter.delete :cov

    begin
      # turn on test coverage
      require 'simplecov'

      SimpleCov.start do
        enable_coverage :branch
        add_filter '/spec/'
        add_filter '/lib/office/word'
        add_filter '/lib/office/excel/sheet_data.rb'
        add_filter '/lib/office/logger.rb'
        add_filter '/lib/office/package.rb'
        add_filter '/lib/office/parts.rb'
        add_filter 'lib/office/constants.rb'
        add_filter 'lib/office/errors.rb'
        add_filter 'lib/office_docs.rb'

        add_group 'Excel', 'lib/office/excel'
        add_group 'Word', 'lib/office/word'
      end

      # output for editor highlighting
      # require 'simplecov-json'
      # SimpleCov.formatter = SimpleCov::Formatter::JSONFormatter
    rescue LoadError
      # do nothing - simplecov not installed
    end
  end
end

require_relative 'modules.rb'

# implement minimal minitest support for migrated tests
class RSpec::Core::ExampleGroup
  def assert_equal a, b
    a.should == b
  end

  def assert pred
    pred.should be_truthy
  end
end

