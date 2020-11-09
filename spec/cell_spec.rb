require 'nokogiri'
require 'date'

require_relative '../lib/office/excel/cell.rb'
require_relative '../lib/office/excel.rb'

require_relative 'spec_helper'

# copy of the minitest test cases, because specs are easier to zero in on
describe Office::Cell do
  class MockXfStyle
    attr_accessor :number_format_id
    attr_accessor :apply_number_format
  end

  class MockStyleSheet
    def initialize
      @ary = []
    end

    attr_reader :ary

    def xf_by_id id
      @ary[id.to_i]
    end
  end

  let :book            do Office::ExcelWorkbook.blank_workbook end
  let :cell            do Office::Cell.new cell_node, nil, styles end
  let :date_epoch      do Date.new(1900, 1, 1) - 2 end
  let :empty_cell_node do node = doc.create_element ?c, r: 'G7'; doc.children.last << node; node end
  let :string_table    do end

  let :doc do
    Nokogiri::XML::Builder.new do |bld|
      bld.root **sheet_namespaces
    end.doc
  end

  let :sheet_namespaces do
    {
      'xmlns' => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      'xmlns:mc' => "http://schemas.openxmlformats.org/markup-compatibility/2006",
      'xmlns:x14ac' => "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
      'xmlns:xr' => "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
      'xmlns:xr2' => "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
      'xmlns:xr3' => "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    }
  end

  let :styles do
    style = MockXfStyle.new
    style.number_format_id = 15
    style.apply_number_format = '1'

    styles = MockStyleSheet.new
    styles.ary[15] = style
    styles
  end

  describe 'date' do
    let :stored_value do Date.today end
    let :cell_node do Office::CellNodes.build_c_node empty_cell_node, stored_value end

    it '#value' do
      day_delta = Date.today - date_epoch
      cell.value.should == day_delta.to_i.to_s
    end

    it '#formatted_value' do
      cell.formatted_value.should == Date.today
    end
  end

  describe 'shared string' do
    it '#value'
    it '#formatted_value'
  end

  describe 'inline string' do
    let :cell_node do Office::CellNodes.build_c_node empty_cell_node, 'Inline String' end

    it '#value' do
      cell.value.should == 'Inline String'
    end

    it '#formatted_value' do
      cell.formatted_value.should == 'Inline String'
    end

    it 'ignores whitespace'
    it 'text runs'
  end

  describe 'integer' do
    let :stored_value do 360 end
    let :cell_node do Office::CellNodes.build_c_node empty_cell_node, stored_value end

    it '#value' do
      cell.value.should == stored_value.to_s
    end

    it '#formatted_value' do
      Integer(cell.formatted_value).should == Integer(stored_value)
    end
  end

  describe 'float' do
    let :stored_value do Math::E end
    let :cell_node do Office::CellNodes.build_c_node empty_cell_node, stored_value end

    it '#value' do
      cell.value.should == stored_value.to_s
    end

    it '#formatted_value' do
      Float(cell.formatted_value).should == Float(stored_value)
    end
  end

  xdescribe 'anything template' do
    let :stored_value do Object.new end
    let :cell_node do Office::CellNodes.build_c_node empty_cell_node, stored_value end

    it '#value' do
      cell.value.should == stored_value.to_s
    end

    it '#formatted_value' do
      cell.formatted_value.should == stored_value
    end
  end

  xdescribe 'text runs' do
    it '#value' do

    end

    it '#formatted_value' do

    end
  end

  describe '#placeholder' do
    it 'nil for no placeholder'
    it 'reset by invalidate'
    it '{{placeholder}}'
    it 'string'
    it 'inline string'
    it 'text runs'
  end
end
