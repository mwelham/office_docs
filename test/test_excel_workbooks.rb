require 'test/unit'
require 'date'
require 'office_docs'

class ExcelWorkbooksTest < Test::Unit::TestCase
  SIMPLE_TEST_WORKBOOK_PATH = File.join(File.dirname(__FILE__), 'content', 'simple_test.xlsx')
  LARGE_TEST_WORKBOOK_PATH = File.join(File.dirname(__FILE__), 'content', 'large_test.xlsx')

  def test_parse_simple_workbook
    book = Office::ExcelWorkbook.new(SIMPLE_TEST_WORKBOOK_PATH)
  end
  
  def test_blank_workbook
    book = Office::ExcelWorkbook.blank_workbook
  end

  def test_simple_csv_export
    book = Office::ExcelWorkbook.new(SIMPLE_TEST_WORKBOOK_PATH)
    assert_equal book.sheets.first.to_csv, "Heading A,Heading B,Heading C\nAlpha,,\nBravo,123,\n,,a;b;c;d\n"
    assert_equal book.sheets.first.to_csv(';'), "Heading A;Heading B;Heading C\nAlpha;;\nBravo;123;\n;;'a;b;c;d'\n"
  end

  def test_parse_large_workbook
    book = Office::ExcelWorkbook.new(LARGE_TEST_WORKBOOK_PATH)
    assert_equal book.sheets.length, 2
    assert book.sheets.first.to_csv.length > 1000
    assert book.sheets.last.to_csv.length > 1000
  end
end
