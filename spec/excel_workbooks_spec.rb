require 'office_docs'
require_relative 'spec_helper'

# copy of the minitest test cases, because specs are easier to zero in on
describe 'ExcelWorkbooksTest' do
  SIMPLE_TEST_WORKBOOK_PATH = File.join __dir__, '/../test/content/simple_test.xlsx'
  LARGE_TEST_WORKBOOK_PATH = File.join __dir__, '/../test/content/large_test.xlsx'
  OPENXML_TEST_WORKBOOK_PATH = File.join __dir__, '/../test/content/openxml_generated.xlsx'

  it :test_parse_simple_workbook do
    Office::ExcelWorkbook.new(SIMPLE_TEST_WORKBOOK_PATH)
  end

  it :test_blank_workbook do
    Office::ExcelWorkbook.blank_workbook
  end

  it :test_simple_csv_export do
    book = Office::ExcelWorkbook.new(SIMPLE_TEST_WORKBOOK_PATH)
    assert_equal book.sheets.first.to_csv, "Heading A,Heading B,Heading C\nAlpha,,\nBravo,123,\n,,a;b;c;d\n"
    assert_equal book.sheets.first.to_csv(';'), "Heading A;Heading B;Heading C\nAlpha;;\nBravo;123;\n;;'a;b;c;d'\n"
  end

  it :test_parse_large_workbook do
    book = Office::ExcelWorkbook.new(LARGE_TEST_WORKBOOK_PATH)
    assert_equal book.sheets.length, 2
    assert book.sheets.first.to_csv.length > 1000
    assert book.sheets.last.to_csv.length > 1000
  end

  it :test_create_workbook do
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

    expected_csv = <<~EOC
      Name,Age,Favorite Virus,Trustworthiness,Spirit Animal
      Alfred,45,Marburg,2.54,
      Carry,6,Measles,0.09,
      Mitch,23,Yellow fever,77,
      Brenda,99,Coxsackie,7.2,Hedgehog
      Greg,345,Rinderpest,-3.1,Possum
      Nathan,23,Hepatitis C,1.3,
      Wilma,21,Canine distemper,8.89,Crocodylocapillaria longiovata
      Arnie,1,Corona,0.0012,Careless Honey Badger
      Phil,0,Dengue,34.5,
    EOC

    assert_equal sheet.to_csv, expected_csv

    file = Tempfile.new('test_create_workbook')
    file.close
    filename = file.path

    book.save(filename)
    assert_equal Office::ExcelWorkbook.new(filename).sheets.first.to_csv, expected_csv

    file.delete
  end

  it :test_from_data do
    book_1 = nil
    File.open(SIMPLE_TEST_WORKBOOK_PATH) { |f| book_1 = Office::ExcelWorkbook.from_data(f.read) }
    book_2 = Office::ExcelWorkbook.new(SIMPLE_TEST_WORKBOOK_PATH)
    assert_equal book_1.sheets.first.to_csv, book_2.sheets.first.to_csv
  end

  it :test_to_data do
    data = Office::ExcelWorkbook.new(SIMPLE_TEST_WORKBOOK_PATH).to_data
    assert !data.nil?
    assert data.length > 0

    book_1 = Office::ExcelWorkbook.from_data(data)
    book_2 = Office::ExcelWorkbook.new(SIMPLE_TEST_WORKBOOK_PATH)
    assert_equal book_1.sheets.first.to_csv, book_2.sheets.first.to_csv
  end

  it :test_sheet_creation do
    book = Office::ExcelWorkbook.blank_workbook
    assert_equal book.sheets.count, 1

    book.add_sheet('Alpha')
    assert_equal book.sheets.count, 2
    book.add_sheet('Bravo')
    assert_equal book.sheets.count, 3
    book.add_sheet('Charlie')
    assert_equal book.sheets.count, 4

    assert_equal book.sheets[1].name, 'Alpha'
    assert_equal book.sheets[2].name, 'Bravo'
    assert_equal book.sheets[3].name, 'Charlie'

    file = Tempfile.new('test_sheet_creation')
    file.close
    filename = file.path
    book.save(filename)

    saved_book = Office::ExcelWorkbook.new(filename)
    assert_equal saved_book.sheets.count, 4
    assert_equal saved_book.sheets[1].name, 'Alpha'
    assert_equal saved_book.sheets[2].name, 'Bravo'
    assert_equal saved_book.sheets[3].name, 'Charlie'
  end

  it :test_sheet_removal do
    book = Office::ExcelWorkbook.blank_workbook
    assert_equal book.sheets.count, 1
    sheet_1 = book.sheets.first
    initial_part_count = book.get_part_names.count

    sheet_2 = book.add_sheet('Another Sheet')
    assert_equal book.sheets.count, 2
    assert_equal book.get_part_names.count, initial_part_count + 1

    sheet_3_name = 'And Another Sheet'
    sheet_3 = book.add_sheet(sheet_3_name)
    assert_equal book.sheets.count, 3
    assert_equal book.get_part_names.count, initial_part_count + 2

    assert_equal sheet_1, book.find_sheet_by_name(sheet_1.name)
    book.remove_sheet(sheet_1)
    assert_equal book.sheets.count, 2
    assert_equal book.get_part_names.count, initial_part_count + 1

    book.remove_sheet(book.find_sheet_by_name(sheet_3_name))
    assert_equal book.sheets.count, 1
    assert_equal book.sheets.first, sheet_2
    assert_equal book.get_part_names.count, initial_part_count
  end

  it :test_parsing_openxml_generated do
    book = Office::ExcelWorkbook.new(OPENXML_TEST_WORKBOOK_PATH)
    assert_equal book.sheets.count, 1
    assert_equal book.sheets.first.name, "Sheet 1"
    assert_equal book.sheets.first.sheet_data.rows.count, 3
    assert_equal book.sheets.first.sheet_data.rows[0].cells.count, 2
    assert_equal book.sheets.first.sheet_data.rows[1].cells.count, 2
    assert_equal book.sheets.first.sheet_data.rows[2].cells.count, 2
    assert_equal book.sheets.first.sheet_data.rows[0].cells[0].value, "Person"
    assert_equal book.sheets.first.sheet_data.rows[0].cells[1].value, "Age"
    assert_equal book.sheets.first.sheet_data.rows[1].cells[0].value, "Person 0001"
    assert_equal book.sheets.first.sheet_data.rows[1].cells[1].value, "20"
    assert_equal book.sheets.first.sheet_data.rows[2].cells[0].value, "Person 0002"
    assert_equal book.sheets.first.sheet_data.rows[2].cells[1].value, "20"
  end

  it 'manipulates a cell' do
    book = Office::ExcelWorkbook.blank_workbook
    sheet = book.sheets.first
    cell = sheet[100,100]
    cell.should be_a(Office::LazyCell)
  end
end
