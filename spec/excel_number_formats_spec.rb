require 'spec_helper'
require 'office_docs'

describe :ExcelNumberFormatsTest do
  SIMPLE_DATA_TYPES_WORKBOOK_PATH = File.join(__dir__, '..', 'test', 'content', 'simple_data_types.xlsx')

  let :book do Office::ExcelWorkbook.new SIMPLE_DATA_TYPES_WORKBOOK_PATH end
  let :sheet do book.sheets.first end
  let :rows do sheet.sheet_data.rows end

  it :test_data_types_parsing do
    assert_equal rows[1].cells[0].formatted_value, Date.new(2001, 1, 1)
    assert_equal rows[2].cells[0].formatted_value, Date.new(2002, 2, 2)
    assert_equal rows[3].cells[0].formatted_value, Date.new(2003, 3, 3)
    assert_equal rows[4].cells[0].formatted_value, Date.new(2004, 4, 4)
    assert_equal rows[5].cells[0].formatted_value, Date.new(2005, 5, 5)

    assert_equal rows[1].cells[1].formatted_value.strftime('%H%M%S'), '010101'
    assert_equal rows[2].cells[1].formatted_value.strftime('%H%M%S'), '020202'
    assert_equal rows[3].cells[1].formatted_value.strftime('%H%M%S'), '030303'
    assert_equal rows[4].cells[1].formatted_value.strftime('%H%M%S'), '040404'
    assert_equal rows[5].cells[1].formatted_value.strftime('%H%M%S'), '050505'


    assert ((rows[1].cells[2].formatted_value - DateTime.new(2001, 1, 1, 1, 1, 1)) * 24 * 60 * 60).abs < 1
    assert ((rows[2].cells[2].formatted_value - DateTime.new(2002, 2, 2, 2, 2, 2)) * 24 * 60 * 60).abs < 1
    assert ((rows[3].cells[2].formatted_value - DateTime.new(2003, 3, 3, 3, 3, 3)) * 24 * 60 * 60).abs < 1
    assert ((rows[4].cells[2].formatted_value - DateTime.new(2004, 4, 4, 4, 4, 4)) * 24 * 60 * 60).abs < 1
    assert ((rows[5].cells[2].formatted_value - DateTime.new(2005, 5, 5, 5, 5, 5)) * 24 * 60 * 60).abs < 1

    assert_equal rows[1].cells[3].formatted_value, 'One'
    assert_equal rows[2].cells[3].formatted_value, 'Two'
    assert_equal rows[3].cells[3].formatted_value, 'Three'
    assert_equal rows[4].cells[3].formatted_value, 'Four'
    assert_equal rows[5].cells[3].formatted_value, 'Five'

    assert (rows[1].cells[4].formatted_value - 1.1).abs < 0.00001
    assert (rows[2].cells[4].formatted_value - 2.2).abs < 0.00001
    assert (rows[3].cells[4].formatted_value - 3.3).abs < 0.00001
    assert (rows[4].cells[4].formatted_value - 4.4).abs < 0.00001
    assert (rows[5].cells[4].formatted_value - 5.5).abs < 0.00001
  end

  it '2d access' do
    book = Office::ExcelWorkbook.new(SIMPLE_DATA_TYPES_WORKBOOK_PATH)
    sheet = book.sheets.first

    sheet[0,1].formatted_value.should == Date.new(2001, 1, 1)
    sheet[4,5].formatted_value.should == 5.5
  end
end
