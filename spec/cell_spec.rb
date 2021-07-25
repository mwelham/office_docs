require_relative 'spec_helper'
require_relative 'version_compatibility'

require 'nokogiri'
require 'date'

require_relative '../lib/office/excel/cell.rb'
require_relative '../lib/office/excel.rb'
require_relative '../lib/office/nokogiri_extensions.rb'

describe Office::Cell do
  using VersionCompatibility

  # Need a bunch of mockup xml stuff here because it's PITA to have a whole
  # workbook just to test cell functionality
  class MockXfStyle
    attr_accessor :number_format_id
    attr_accessor :apply_number_format
    def ignore_number_format?; apply_number_format.to_i != 1 end
  end

  class MockStyleSheet
    def initialize
      @ary = []
    end

    attr_reader :ary

    def xf_by_index index
      @ary[index.to_i]
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
      #  ≥ruby-2.7 does not need the transform_keys
      bld.root **sheet_namespaces.transform_keys(&:to_sym)
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
    # In real xlsx, the style indexing isn't one-to-one like this.
    # But here we're mapping directly to the number format ids for easier thinking about.
    # Hopefully that doesn't break anything.
    MockStyleSheet.new.tap do |styles|
      # for Integer
      styles.ary[1] = MockXfStyle.new.tap do |style|
        style.number_format_id = 1
        style.apply_number_format = '1'
      end

      # for Date
      styles.ary[14] = MockXfStyle.new.tap do |style|
        style.number_format_id = 14
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

  describe 'build_c_node' do
    let :cell_node do Office::CellNodes.build_c_node empty_cell_node, stored_value, styles: styles end

    describe 'error' do
      let :stored_value do Object.new end
      it 'raises with Object' do
        ->{cell_node}.should raise_error(Office::TypeError)
      end
    end

    describe Date do
      let :stored_value do Date.today end

      it '#value' do
        day_delta = stored_value - Office::CellNodes::DATE_EPOCH
        cell.value.should == Integer(day_delta).to_s
      end

      it '#formatted_value' do
        cell.formatted_value.should == Date.today
      end
    end

    describe DateTime do
      let :stored_value do DateTime.now end

      it '#value' do
        day_delta = Float stored_value.to_time.floor.to_datetime - (Office::CellNodes::DATE_TIME_EPOCH - Office::CellNodes::UTC_OFFSET_HOURS)
        cell.value.should == day_delta.to_s
      end

      it '#formatted_value' do
        # precision is 1/86400 = 1 / 24*60*60
        cell.formatted_value.should == stored_value.to_time.floor.to_datetime
      end
    end

    # pretty much same as DateTime
    describe Time do
      let :stored_value do Time.now end

      it '#value' do
        day_delta = Float stored_value.floor.to_datetime - (Office::CellNodes::DATE_TIME_EPOCH - Office::CellNodes::UTC_OFFSET_HOURS)
        cell.value.should == day_delta.to_s
      end

      it '#formatted_value' do
        # precision is 1/86400 = 1 / 24*60*60
        cell.formatted_value.should == stored_value.floor
      end
    end

    describe 'shared string' do
      it '#value'
      it '#formatted_value'
    end

    describe String do
      let :stored_value do 'Inline String' end

      it '#value' do
        cell.value.should == stored_value.to_s
      end

      it '#formatted_value' do
        cell.formatted_value.should == stored_value
      end

      it 'ignores whitespace'
      it 'text runs'
    end

    describe Integer do
      let :stored_value do 360 end

      it '#value' do
        cell.value.should == stored_value.to_s
      end

      it '#formatted_value' do
        Integer(cell.formatted_value).should == Integer(stored_value)
      end
    end

    describe Float do
      let :stored_value do Math::E end

      it '#value' do
        cell.value.should == stored_value.to_s
      end

      it '#formatted_value' do
        Float(cell.formatted_value).should == Float(stored_value)
      end
    end

    describe 'IsoTime' do
      let :stored_value do Office::IsoTime.new Time.now end

      it '#value' do
        cell.value.should == stored_value.iso8601
      end

      it '#formatted_value' do
        cell.formatted_value.should == stored_value.time.floor
      end
    end

    xdescribe 'anything template' do
      let :stored_value do Object.new end

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

  describe Office::Cell::Placeholder do
    MockCell = Struct.new :value

    describe 'of_cuddled' do
      it '{{placeholder}}' do
        placeholder_str = '{{placeholder}}'
        plch = described_class.of_cuddled MockCell.new(placeholder_str), placeholder_str

        plch.to_s.should == placeholder_str
        plch.start.should == 0
        plch.length.should == 15
      end

      it '\n\n{{placeholder}}\n\n' do
        placeholder_str = '\n\n{{placeholder}}\n\n'
        plch = described_class.of_cuddled MockCell.new(placeholder_str), placeholder_str

        plch.to_s.should == '{{placeholder}}'
        plch.start.should == 4
        plch.length.should == 15
      end

      it '\n\n  {{placeholder}} and some more words\n\n' do
        placeholder_str = '\n\n  {{placeholder}} and some more words\n\n'
        plch = described_class.of_cuddled MockCell.new(placeholder_str), placeholder_str

        plch.to_s.should == '{{placeholder}}'
        plch.start.should == 6
        plch.length.should == 15
      end

      it 'plain text' do
        described_class.of_cuddled(MockCell.new('plain_text'), 'plain_text').should be_nil
      end

      it '{{broken placeholder}' do
        described_class.of_cuddled(MockCell.new('{{broken placeholder}'), '{{broken placeholder}').should be_nil
      end

      it '{{broken placeholder' do
        described_class.of_cuddled(MockCell.new('{{broken placeholder'), '{{broken placeholder').should be_nil
      end

      it 'nil for no placeholder' do
        described_class.of_cuddled(MockCell.new(nil),nil).should be_nil
      end
    end
  end

  describe '#value=' do
    xit 'shared String'

    it 'Date' do
      cell.value.should be_nil
      date_value = Date.today
      xlsx_repr = Integer date_value - Office::CellNodes::DATE_EPOCH
      cell.value = date_value
      cell.node.to_xml.should == <<~EOX.chomp
      <c r="G7" t="n" s="14">
        <v>#{xlsx_repr}</v>
      </c>
      EOX
      Office::Cell.new(cell.node, nil, styles).to_ruby.should be_a(Date)
    end

    it 'Time' do
      cell.value.should be_nil
      time_value = Time.now
      cell.value = time_value
      xlsx_repr = Float time_value.floor.to_datetime - (Office::CellNodes::DATE_TIME_EPOCH - Office::CellNodes::UTC_OFFSET_HOURS)
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
      cell.value = date_value
      xlsx_repr = Float date_value.to_time.floor.to_datetime - (Office::CellNodes::DATE_TIME_EPOCH - Office::CellNodes::UTC_OFFSET_HOURS)
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
