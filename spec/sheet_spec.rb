require_relative '../lib/office/excel'
require_relative '../lib/office/constants'
require_relative '../lib/office/nokogiri_extensions'

require_relative 'spec_helper'
require_relative 'xml_fixtures'

# copy of the minitest test cases, because specs are easier to zero in on
describe Office::Sheet do
  include XmlFixtures
  include ReloadWorkbook

  let :book do
    Office::ExcelWorkbook.new File.join(__dir__, '/../test/content/simple-placeholders.xlsx')
  end

  let :sheet do book.sheets.first end

  describe 'placeholders' do
    let :placeholders do
      ["{{horizontal}}", "{{manufacturer}}", "{{yes}}", "{{shop_or_serial}}", "{{model_number}}", "{{vertical}}", "{{no}}", "{{gpm}}", "{{rated_head_foot_psi}}", "{{net_psi}}", "{{rated_rpm}}", "very {{important}} thing", "{{broken_place}}", "{{streams|tabular}}"]
    end

    it 'finds placeholders - generic' do
      place_cells = sheet.each_cell.filter(&:placeholder)
      place_cells.map(&:value).should == placeholders
    end

    it '#each_placeholder' do
      sheet.each_placeholder.map(&:value).should == placeholders
    end
  end

  describe 'cell operations' do
    let :book do
      Office::ExcelWorkbook.new BookFiles::SIMPLE_TEST
    end

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
    let :book do
      Office::ExcelWorkbook.new File.join(__dir__, '/../test/content/simple-placeholders.xlsx')
    end

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
        range = Office::Range.new 'A5:I17'
        sheet.delete_rows range
        sheet.dimension = sheet.calculate_dimension
        sheet.sheet_data.rows.count.should == 5
        # TODO verify that both cells and rows are renumbered correctly

        reload_workbook sheet.workbook, 'delete.xlsx' do |book|
          # `localc --nologo #{book.filename}`
        end
      end
    end

    describe 'tabular data' do
      let :dataset do
        [
        %i[rpm discharge suction net no size pitot gpm percent voltage amp],
        [1791,110,10,100,0,'Churn',nil,0,0,"477,476,477","56,55,56"],
        [1789,106,5,101,3,2.5,7.6,1500,100,"479,475,475","90,86,91"],
        [1777,96,6,91,3,2.5,17,2250,150,"477,475,474","118,115,118"],
        [1770,92,5,87,3,8.5,23,3000,170,"452,454,424","128,135,132"],
        [1762,83,2,82,3,6.1,23,3750,80,"454,424,403","124,128,135"],
        [1751,65,1,79,3,4.5,23,4500,50,"424,403,389","121,124,128"],
        ]
      end

      it 'insert rows with tabular data' do
        placeholder_cell = sheet.each_placeholder.find{|cell| cell.value =~ /tabular/}
        placeholder_cell.should_not be_nil

        # make the title range into normal cells
        title_range = sheet.merge_ranges.find{|range| range.cover? placeholder_cell.location}
        title_range.should_not be_nil
        sheet.delete_merge_range title_range

        first_row, *records = dataset

        # overwrite header row
        first_row.each_with_index do |header_value,colix|
          insert_location = placeholder_cell.location + [colix, 0]
          sheet[insert_location].value = header_value.to_s
        end

        # insert blank rows for data, from row after first_row
        insert_range = Office::Range.new(placeholder_cell.location + [0,1], placeholder_cell.location + [records.first.size,records.size])
        inserted_rows = sheet.insert_rows(insert_range)

        # TODO optimise this by creating a map from inserted_rows
        # and/or allowing Cell to work on a fragment as well as on the full sheet.
        # and/or allowing the Sheet#[] to work on a nodeset of rows.
        # which means having indexing be more flexible than an array of Row instances.
        records.each_with_index do |data_row, rowix|
          data_row.each_with_index do |val, colix|
            location = insert_range.top_left + [colix, rowix]
            sheet[location].value = val
          end
        end

        # update sheet dimension
        sheet.dimension = sheet.calculate_dimension

        reload_workbook sheet.workbook, 'insert.xlsx' do |book|
          # `localc --nologo #{book.filename}`
        end
      end

      it 'replaces cells with tabular data' do
        # TODO put most of this in Sheet#replace_tabular
        placeholder_cell = sheet.each_placeholder.find{|cell| cell.value =~ /tabular/}
        placeholder_cell.should_not be_nil

        range = sheet.merge_ranges.find{|range| range.cover? placeholder_cell.location}
        range.should_not be_nil

        first_row, *records = dataset

        # set first_row
        first_row.each_with_index do |header_value,colix|
          insert_location = placeholder_cell.location + [colix, 0]
          sheet[insert_location].value = header_value.to_s
        end

        # overwrite data values after header
        record_location = placeholder_cell.location + [0,1]
        records.each_with_index do |data_row, rowix|
          # check that range matches data
          data_row.size <= range.width or raise "data too long for #{range}: (#{data_row.size})#{data_row.inspect}"
          data_row.each_with_index do |val, colix|
            location = record_location + [colix, rowix]
            # TODO insert second and subsequent rows
            # TODO what's the difference between .value = and []= ? The latter would not require LazyCell
            # also [] and []= could apply to rows, but we're kinda using sheet.sheet_data.rows for that
            sheet[location].value = val
          end
        end

        # make the title range into normal cells
        title_range = sheet.merge_ranges.find{|range| range.cover? placeholder_cell.location}
        title_range.should_not be_nil
        sheet.delete_merge_range title_range

        # TODO update sheet range. Need to find largest cell reference number. Oof.
        # possibly grab current range, extend by loc_track, but then still need to check for rows bumped by insertion
        # unless every row insertion / row deletion / lazy cell insertion tracks the current range.
        # Hmmm. Just do it brute-force for now.
        sheet.dimension.should == 'A5:K22'
        sheet.calculate_dimension.should == 'A5:K24'
        sheet.dimension = sheet.calculate_dimension
        sheet.dimension.should == 'A5:K24'

        reload_workbook sheet.workbook, 'overwrite.xlsx' do |book|
          # `localc --nologo #{book.filename}`
        end
      end

      it "appends rows to end if they don't exist"
      it "inserts rows if necessary"
      it "creates new empty row if it doesn't exist"
    end
  end
end
