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
  
  def test_create_workbook
    book = Office::ExcelWorkbook.blank_workbook
    sheet = book.sheets.first

    sheet.add_row [ "Name", "Age", "Favorite Virus", "Trustworthiness", "Spirit Animal" ]
    sheet.add_row [ "Alfred", 45, "Marburg", 2.54, nil ]
    sheet.add_row [ "Carry", 6, "Measles", 0.09, "" ]
    sheet.add_row [ "Mitch", 23, "Yellow fever", 77 ]
    sheet.add_row [ "Brenda", 99, "Coxsackie", 7.2, "Hedgehog" ]
    sheet.add_row [ "Greg", 345, "Rinderpest", -3.1, "Possum" ]
    sheet.add_row [ "Nathan", 23, "Hepatitis C", 1.3, "" ]
    sheet.add_row [ "Wilma", 21, "Canine distemper", 8.89, "Crocodylocapillaria longiovata" ]
    sheet.add_row [ "Arnie", 1, "Corona", 0.0012, "Careless Honey Badger" ]
    sheet.add_row [ "Phil", 0, "Dengue", 34.5 ]

    expected_csv = "Name,Age,Favorite Virus,Trustworthiness,Spirit Animal\n" +
      "Alfred,45,Marburg,2.54,\n" +
      "Carry,6,Measles,0.09,\n" +
      "Mitch,23,Yellow fever,77,\n" +
      "Brenda,99,Coxsackie,7.2,Hedgehog\n" +
      "Greg,345,Rinderpest,-3.1,Possum\n" +
      "Nathan,23,Hepatitis C,1.3,\n" +
      "Wilma,21,Canine distemper,8.89,Crocodylocapillaria longiovata\n" +
      "Arnie,1,Corona,0.0012,Careless Honey Badger\n" +
      "Phil,0,Dengue,34.5,\n"
    
    assert_equal sheet.to_csv, expected_csv
    
    file = Tempfile.new('test_create_workbook')
    file.close
    filename = file.path
    
    book.save(filename)
    assert_equal Office::ExcelWorkbook.new(filename).sheets.first.to_csv, expected_csv
    
    file.delete
  end
end
