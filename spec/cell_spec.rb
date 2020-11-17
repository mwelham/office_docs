require 'nokogiri'
require 'date'

require_relative '../lib/office/excel/cell.rb'
require_relative '../lib/office/excel.rb'
require_relative '../lib/office/nokogiri_extensions.rb'

require_relative 'spec_helper'

describe Office::Cell do
  # Need a bunch of mockup xml stuff here because it's PITA to have a whole
  # workbook just to test cell functionality
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

    def index_of_xf num_fmt_id
      # slow but it's for testing
      @ary.index{|xf| xf&.number_format_id == num_fmt_id}
    end
  end

  let :cell_node       do Office::CellNodes.build_c_node empty_cell_node, nil, styles: styles end
  let :cell            do Office::Cell.new cell_node, nil, styles end
  let :location        do 'G7' end
  let :empty_cell_node do node = doc.create_element ?c, r: location; doc.children.last << node; node end

  # need a minimal doc as backing for Cell instances
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
    # In real xlsx, the style indexing doesn't work like this.
    # But here we're mapping directly to the number format ids for easier thinking about.
    # Hopefully that doesn't break anything.
    MockStyleSheet.new.tap do |styles|
      # for Date
      styles.ary[1] = MockXfStyle.new.tap do |style|
        style.number_format_id = 1
        style.apply_number_format = '1'
      end

      # for Date
      styles.ary[15] = MockXfStyle.new.tap do |style|
        style.number_format_id = 15
        style.apply_number_format = '1'
      end

      # for DateTime
      styles.ary[22] = MockXfStyle.new.tap do |style|
        style.number_format_id = 22
        style.apply_number_format = '1'
      end

      # for Time
      styles.ary[21] = MockXfStyle.new.tap do |style|
        style.number_format_id = 21
        style.apply_number_format = '1'
      end
    end
  end

  describe 'Office::CellNodes.build_c_node' do
    describe 'date' do
      let :stored_value do Date.today end
      let :cell_node do Office::CellNodes.build_c_node empty_cell_node, stored_value, styles: styles end

      it '#value' do
        day_delta = Date.today - Office::CellNodes::DATE_EPOCH
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
      let :cell_node do Office::CellNodes.build_c_node empty_cell_node, 'Inline String', styles: styles end

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
      let :cell_node do Office::CellNodes.build_c_node empty_cell_node, stored_value, styles: styles end

      it '#value' do
        cell.value.should == stored_value.to_s
      end

      it '#formatted_value' do
        Integer(cell.formatted_value).should == Integer(stored_value)
      end
    end

    describe 'float' do
      let :stored_value do Math::E end
      let :cell_node do Office::CellNodes.build_c_node empty_cell_node, stored_value, styles: styles end

      it '#value' do
        cell.value.should == stored_value.to_s
      end

      it '#formatted_value' do
        Float(cell.formatted_value).should == Float(stored_value)
      end
    end

    xdescribe 'anything template' do
      let :stored_value do Object.new end
      let :cell_node do Office::CellNodes.build_c_node empty_cell_node, stored_value, styles: styles end

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

    describe 'without style' do
      let :cell_node do
        # build the node outside of the normal Cell#value= code, ie no type conversions
        # <c r="B3">
        #   <v>7919</v>
        # </c>
        c_node = empty_cell_node.document.build_element ?c, r: 'F1' do |bld|
          bld.v cell_value
        end
      end

      describe 'Integer' do
        let :cell_value do "7919" end
        it {cell.formatted_value.should == 7919}
      end

      describe 'Float' do
        let :cell_value do "7.919" end
        it {cell.formatted_value.should == 7.919}
      end

      describe 'String' do
        let :cell_value do "0.0ops" end
        it {cell.formatted_value.should == "0.0ops"}
      end
    end
  end

  xdescribe '#placeholder' do
    it 'nil for no placeholder'
    it 'reset by invalidate'
    it '{{placeholder}}'
    it 'string'
    it 'inline string'
    it 'text runs'
  end

  describe '#value=' do
    xit 'shared String'

    it 'Date' do
      cell.value.should be_nil
      date_value = Date.today
      xlsx_repr = Integer date_value - Office::CellNodes::DATE_EPOCH
      cell.value = date_value
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7" t="n" s="15">
        <v>#{xlsx_repr}</v>
      </c>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should be_a(Date)
    end

    it 'Time' do
      cell.value.should be_nil
      time_value = Time.now
      xlsx_repr = Float time_value.to_datetime - Office::CellNodes::DATE_TIME_EPOCH
      cell.value = time_value
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7" t="n" s="21">
        <v>#{xlsx_repr}</v>
      </c>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should be_a(Time)
    end

    it 'DateTime' do
      cell.value.should be_nil
      date_value = DateTime.now
      xlsx_repr = Float date_value - Office::CellNodes::DATE_TIME_EPOCH
      cell.value = date_value
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7" t="n" s="22">
        <v>#{xlsx_repr}</v>
      </c>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should be_a(DateTime)
    end

    it 'nil' do
      cell.value.should be_nil
      cell.value = nil
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7"/>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should be_nil
    end

    it 'boolean true' do
      cell.value.should be_nil
      cell.value = true
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7" t="b" s="0">
        <v>1</v>
      </c>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should == true
    end

    it 'boolean false' do
      cell.value.should be_nil
      cell.value = false
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7" t="b" s="0">
        <v>0</v>
      </c>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should == false
    end

    it 'inline String' do
      cell.value.should be_nil
      cell.value = 'Hello Darlink'
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7" t="inlineStr" s="0">
        <is>
          <t>Hello Darlink</t>
        </is>
      </c>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should be_a(String)
    end

    it 'Integer' do
      cell.value.should be_nil
      random_int = rand -18446744073709551615 .. 18446744073709551615
      cell.value = random_int
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7" t="n" s="1">
        <v>#{random_int}</v>
      </c>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should be_a(Integer)
    end

    it 'Float' do
      cell.value.should be_nil
      random_float = rand Float::MIN .. Float::MAX
      cell.value = random_float
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7" t="n" s="0">
        <v>#{random_float}</v>
      </c>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should be_a(Float)
    end
  end
end
