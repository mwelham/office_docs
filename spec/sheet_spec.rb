require_relative 'spec_helper'

require_relative '../lib/office/excel'
require_relative '../lib/office/constants'
require_relative '../lib/office/nokogiri_extensions'

require_relative 'xml_fixtures'

# copy of the minitest test cases, because specs are easier to zero in on
describe Office::Sheet do
  include XmlFixtures
  include ReloadWorkbook

  let :book do Office::ExcelWorkbook.new BookFiles::SIMPLE_PLACEHOLDERS end
  let :sheet do book.sheets.first end

  describe 'placeholders' do
    let :placeholders do
      ["{{horizontal}}", "{{manufacturer}}", "{{yes}}", "{{shop_or_serial}}", "{{model_number}}", "{{vertical}}", "{{no}}", "{{gpm}}", "{{rated_head_foot_psi}}", "{{net_psi}}", "{{rated_rpm}}", "very {{important}} thing", "{{broken_place}}", "{{streams|tabular}}"]
    end

    it 'finds placeholders - generic' do
      place_cells = sheet.each_cell.select(&:placeholder)
      place_cells.map(&:value).should == placeholders
    end

    it '#each_placeholder' do
      sheet.each_placeholder.map(&:value).should == placeholders
    end
  end

  describe 'cell operations' do
    let :book do Office::ExcelWorkbook.new BookFiles::SIMPLE_TEST end

    it 'replaces one cell' do
      the_value = "This is the new pump"
      cell = sheet.each_cell.first
      cell.value = the_value

      reload_workbook sheet.workbook do |book|
        saved_cell = book.sheets.first.each_cell.first
        saved_cell.value.should == the_value
        saved_cell.node.object_id.should_not == cell.node.object_id
      end
    end

    # TODO probably belongs in its own spec file
    describe Office::LazyCell do
      let :book do Office::ExcelWorkbook.blank_workbook end

      # a cell outside the current (recorded) sheet dimension
      let :outside_loc do sheet.dimension.bot_rite + [10,10] end

      it 'make sense' do
        # get a cell that doesn't exist yet
        lazy_cell = sheet[outside_loc]
        lazy_cell.should be_a(Office::LazyCell)
        lazy_cell.should be_empty
        lazy_cell.value.should be_nil
        lazy_cell.formatted_value.should be_nil
      end

      it 'value= creates row and cell' do
        # assert precondition - there are no rows
        ndst = sheet.sheet_data.node.nspath("~row[@r=#{outside_loc.row_r}]")
        ndst.size.should == 0

        # set value on a cell outside the current sheet dimension
        sheet[outside_loc].value = "Cat on a Warm Dry Chair"
        ndst = sheet.sheet_data.node.nspath("~row[@r=#{outside_loc.row_r}]")
        ndst.size.should == 1

        # find the cell node outside of the location layer
        cell_node = sheet.sheet_data.node.nspath("~row[@r=#{outside_loc.row_r}]/~c").first
        # .text will not work here if the created node is a shared string
        cell_node.text.should == "Cat on a Warm Dry Chair"
      end
    end

    describe '#each_cell' do
      it 'iterates cells' do
        sheet.each_cell.map(&:data_type).should == [:s, :s, :s, :s, :s, nil, :s]
        sheet.each_cell.map(&:formatted_value).should ==["Heading A", "Heading B", "Heading C", "Alpha", "Bravo", 123, "a;b;c;d"]
      end
    end
  end

  describe 'row operations' do
    describe '#insert_rows' do
      it 'from range' do
        sheet.dimension.height.should == 18
        sheet.dimension.width.should == 11

        old_value = sheet['H8'].value

        # insert 5 rows starting from row 8 shifting all others down
        sheet.insert_rows(Office::Location.new('H8') * [1,5])

        sheet.calculate_dimension.height.should == 23
        sheet.dimension.width.should == 11

        sheet['H13'].value.should == old_value
        sheet['H8'].value.should be_nil
      end
    end

    describe '#delete_rows' do
      it 'from range' do
        # delete 12 rows at the top
        delete_range = sheet.dimension.top_left * [1,13]
        deleted_rows = sheet.delete_rows delete_range
        sheet.calculate_dimension.height.should == 5

        # verify that both cells and rows are renumbered correctly
        cells = sheet.each_cell_by_node.sort_by{|cell| cell.location.to_a}
        cells.map{|cell| cell.location.to_s}.should == %w[A5 A6 B5 B6 C5 C6 D5 D6 E5 F5 G5 H5 I5 J5 J8 J9 K5 K7 K8 K9]
      end
    end
  end

  describe 'tabular operations' do
    let :dataset do
      [
      %w[rpm discharge suction net no size pitot gpm percent voltage amp],
      [1791,110,10,100,0,'Churn',nil,0,0,"477,476,477","56,55,56"],
      [1789,106,5,101,3,2.5,7.6,1500,100,"479,475,475","90,86,91"],
      [1777,96,6,91,3,2.5,17,2250,150,"477,475,474","118,115,118"],
      [1770,92,5,87,3,8.5,23,3000,170,"452,454,424","128,135,132"],
      [1762,83,2,82,3,6.1,23,3750,80,"454,424,403","124,128,135"],
      [1751,65,1,79,3,4.5,23,4500,50,"424,403,389","121,124,128"],
      ]
    end

    let :placeholder_cell do sheet.each_placeholder.find{|cell| cell.value =~ /streams\|tabular/} end

    # example of how to use locations, ranges, insert_rows and accept!
    it 'insert rows with tabular data' do
      sheet.dimension.should == sheet.calculate_dimension
      saved_dimension = sheet.dimension

      # verify placeholder
      placeholder_cell.should_not be_nil

      # make the title range into normal cells
      title_range = sheet.merge_ranges.find{|range| range.cover? placeholder_cell.location}
      title_range.should_not be_nil
      sheet.delete_merge_range title_range

      # insert blank rows for data, from row after first_row
      insert_range = (placeholder_cell.location + [0,1]) * [dataset.first.size, dataset.size-1]
      insert_range.to_s.should == 'A19:K24'
      sheet.insert_rows(insert_range)

      # set values of inserted rows from size of dataset
      sheet.accept! placeholder_cell.location, dataset

      # make sure invalidate has marked dimension for recalc
      sheet.dimension.should == sheet.calculate_dimension
      # obviously dimension should change
      saved_dimension.should_not == sheet.dimension

      # reload_workbook sheet.workbook, 'insert.xlsx' do |book| `localc --nologo #{book.filename}` end
      cells = sheet.cells_of (placeholder_cell.location * [dataset.first.size, dataset.size]), &:formatted_value
      cells.should == dataset
    end

    describe '#accept!' do
      it 'sets values' do
        # verify placeholder
        placeholder_cell.should_not be_nil

        range = sheet.accept! placeholder_cell.location, dataset

        # reload_workbook sheet.workbook, 'accept.xlsx' do |book| `localc --nologo #{book.filename}` end

        # gather values changed cells
        values = sheet.cells_of range, &:formatted_value

        values.should == dataset
      end
    end

    describe '#project!' do
      it 'sets values' do
        # verify placeholder
        placeholder_cell.should_not be_nil

        # for project! we have to tell the method which range to fill
        range = placeholder_cell.location * [dataset.first.size, dataset.size]
        sheet.project! range, dataset

        # gather values changed cells
        values = sheet.cells_of range, &:formatted_value

        # reload_workbook sheet.workbook, 'project.xlsx' do |book| `localc --nologo #{book.filename}` end

        values.should == dataset
      end
    end
  end

  describe 'dimensions' do
    describe 'A1 for blank sheet' do
      let :book do Office::ExcelWorkbook.blank_workbook end

      it '#dimension' do
        sheet.send(:dimension_node)[:ref].should == 'A1'
        sheet.dimension.should == Office::Range.new('A1:A1')
      end

      it '#calculate_dimension' do
        sheet.calculate_dimension.should == Office::Range.new('A1:A1')
      end
    end

    it 'saves calculated dimension' do
      saved_dimension = sheet.dimension
      calc_dimension = sheet.calculate_dimension
      saved_dimension.should == calc_dimension

      outside_loc = saved_dimension.bot_rite + [2,2]
      sheet[outside_loc].value = 'outside violin solo'

      reload_workbook book do |book|
        # grab the dimension node before anything else
        book.sheets.first.send(:dimension_node)[:ref].should == Office::Range.new(saved_dimension.top_left, outside_loc).to_s

        book.sheets.first.dimension.should_not == saved_dimension
        book.sheets.first.dimension.should == Office::Range.new(saved_dimension.top_left, outside_loc)
      end
    end
  end

  describe 'range fetching' do
    let :book do Office::ExcelWorkbook.new(BookFiles::LARGE_TEST) end

    it 'cells', performance: true do
      sheet = book.sheets.first
      last_row = sheet.dimension.bot_rite

      require 'benchmark'
      Benchmark.bmbm do |results|
        results.report 'cell_nodes_of' do sheet.cell_nodes_of(range: Office::Range.new('G12:BJ63'), &sheet.method(:cell_of)) end
        results.report 'cells_of' do sheet.cells_of Office::Range.new('G12:BJ63') end
        results.report 'cells_of formatted' do sheet.cells_of Office::Range.new('G12:BJ63'), &:formatted_value end
      end
    end
  end

  describe 'csv' do
    let :book do Office::ExcelWorkbook.new(BookFiles::LARGE_TEST) end
    let :sheet do book.sheets.first end

    it 'comparisons', performance: true do
      require 'benchmark'
      Benchmark.bmbm do |results|
        results.report 'to_excel_csv' do sheet.to_excel_csv end
        results.report 'old_range_to_csv' do sheet.send :old_range_to_csv end
        results.report 'range_to_csv' do sheet.range_to_csv end
        results.report 'cells_of to_csv' do sheet.cells_of(&:formatted_value).to_csv end
      end
    end
  end
end
