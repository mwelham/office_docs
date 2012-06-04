require 'test/unit'
require 'date'
require 'office_docs'

class ExcelWorkbooksTest < Test::Unit::TestCase
  SIMPLE_TEST_WORKBOOK_PATH = File.join(File.dirname(__FILE__), 'content', 'simple_test.xlsx')

  def test_parse_simple_workbook
    book = Office::ExcelWorkbook.new(SIMPLE_TEST_WORKBOOK_PATH)
    book.debug_dump
  end
  
  def test_blank_workbook
    book = Office::ExcelWorkbook.blank_workbook
  end
end
