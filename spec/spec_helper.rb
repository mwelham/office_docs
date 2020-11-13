require 'rspec'

RSpec.configure do |config|
  # enable should syntax
  config.expect_with(:rspec) { |c| c.syntax = :should }

  # allow for --only-failures
  config.example_status_persistence_file_path = "spec/examples.txt"

  # exclude things marked as performance benchmarks
  config.filter_run_excluding performance: true, display_ui: true
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

module BookFiles
  (Pathname(__dir__) + "../test/content").children.each do |path|
    next unless path.to_s.end_with? '.xlsx'
    const_name = path.basename('.xlsx').to_s.gsub(/[[:punct:]]/, ?_).upcase
    const_set const_name, path.realpath.to_s
  end
end

module ReloadWorkbook
  def reload_workbook workbook, filename = nil, &blk
    Dir.mktmpdir do |dir|
      filename = File.join dir, (filename || File.basename(workbook.filename))
      workbook.save filename
      yield Office::ExcelWorkbook.new(filename)
    end
  end
end

