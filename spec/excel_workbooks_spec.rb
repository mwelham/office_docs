require_relative 'spec_helper'
require_relative 'version_compatibility'

require 'csv'
require 'office_docs'

# copy of the minitest test cases, because specs are easier to zero in on
describe Office::ExcelWorkbook do
  include ReloadWorkbook
  using VersionCompatibility

  it :test_parse_simple_workbook do
    Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST)
  end

  it :test_blank_workbook do
    Office::ExcelWorkbook.blank_workbook
  end

  it :test_simple_csv_export do
    book = Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST)
    assert_equal book.sheets.first.to_excel_csv, "Heading A,Heading B,Heading C\nAlpha,,\nBravo,123,\n,,a;b;c;d\n"
    assert_equal book.sheets.first.to_excel_csv(';'), "Heading A;Heading B;Heading C\nAlpha;;\nBravo;123;\n;;'a;b;c;d'\n"
  end

  it :test_parse_large_workbook do
    book = Office::ExcelWorkbook.new(FixtureFiles::Book::LARGE_TEST)
    assert_equal book.sheets.length, 2
    assert book.sheets.first.to_csv.length > 1000
    assert book.sheets.last.to_csv.length > 1000
  end

  it :test_create_workbook do
    values = [
      [ "Name", "Age", "Favorite Virus", "Trustworthiness", "Spirit Animal" ],
      [ "Alfred", 45, "Marburg", 2.54, nil ],
      [ "Carry", 6, "Measles", 0.09, "" ],
      [ "Mitch", 23, "Yellow fever", 77 ],
      [ "Brenda", 99, "Coxsackie", 7.2, "Hedgehog" ],
      [ "Greg", 345, "Rinderpest", -3.1, "Possum" ],
      [ "Nathan", 23, "Hepatitis C", 1.3, "" ],
      [ "Wilma", 21, "Canine distemper", 8.89, "Crocodylocapillaria longiovata" ],
      [ "Arnie", 1, "Corona", 0.0012, "Careless Honey Badger" ],
      [ "Phil", 0, "Dengue", 34.5 ],
    ]
    book = Office::ExcelWorkbook.blank_workbook
    sheet = book.sheets.first
    values.each{|ary| sheet.add_row ary }

    assert_equal sheet.each_cell.map(&:formatted_value), values.flatten

    Dir.mktmpdir do |dir|
      filename = File.join dir, 'previous-blank.xlsx'
      book.save(filename)
      saved_book = Office::ExcelWorkbook.new(filename)
      assert_equal saved_book.sheets.first.each_cell.map(&:formatted_value), values.flatten
    end
  end

  describe 'type-styles' do
    let :book do Office::ExcelWorkbook.blank_workbook end

    let :values do
      [ Date.today, Time.now, DateTime.now, "Stringly-typed", 1.6180339887, 42, true, false, nil, Office::IsoTime.new(Time.now) ]
    end

    # Floating point date/times use .floor to get rid of fractional seconds
    # because xlsx doesn't understand those.
    let :exact_values do
      values.map do |val|
        case val
        when DateTime;        val.to_time.floor.to_datetime
        when Time;            val.floor
        when Date;            val
        when Office::IsoTime; val.time.floor
        else;                 val
        end
      end
    end

    let :row_range do (sheet.dimension.bot_left + [0,1]) * [values.size,1] end

    let :sheet do
      book.sheets.first
    end

    describe 'in worksheet with styles' do
      let :book do Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_DATA_TYPES) end

      # TODO this fails because style entries for dates do not exist, and need to be created
      it 'has correct types' do
        # insert values. Don't use add_rows because that doesn't set type-styles properly
        # one row down and extend to width of values (not width of existing dimension)
        row_range.each_by_row.with_index do |loc,index|
          sheet[loc].value = values[index]
        end

        sheet.invalidate_row_cache

        cells = row_range.each_by_row.map{|loc| sheet[loc]}

        cells.zip(exact_values).each.with_index do |(cell, exact_value), _index|
          cell.formatted_value.should == exact_value
        end
      end
    end

    describe 'in blank worksheet' do
      let :book do Office::ExcelWorkbook.blank_workbook end

      # TODO this fails because style entries for dates do not exist, and need to be created
      it 'has correct types' do
        row_range.each_by_row.with_index do |loc,index|
          sheet[loc].value = values[index]
        end

        sheet.invalidate_row_cache

        cells = row_range.each_by_row.map{|loc| sheet[loc]}
        cells.zip(exact_values) do |(cell, exact_value)|
          cell.formatted_value.should == exact_value
        end
      end
    end

    it 'large numbers and strange types', display_ui: true do
      sheet['A1'].value = "Large Integer"
      sheet['A2'].value = rand -18446744073709551615 .. 18446744073709551615

      sheet['B1'].value = "Large Float"
      sheet['B2'].value = rand Float::MIN .. Float::MAX

      sheet['C1'].value = "d time"
      sheet['C2'].value = Office::IsoTime.new Time.now

      sheet['D1'].value = "normal time"
      sheet['D2'].value = Time.now

      sheet.invalidate_row_cache

      reload_workbook book, 'large_strange_types.xlsx' do |book|
        `localc --nologo #{book.filename}`
      end
    end

    it 'saved file still correct' do
      row_range.each_by_row.with_index do |loc,index|
        sheet[loc].value = values[index]
      end

      sheet.invalidate_row_cache

      reload_workbook book, 'previously_blank' do |book|
        sheet = book.sheets.first
        cells = row_range.each_by_row.map{|loc| sheet[loc]}
        cells.zip(exact_values).each do |cell, exact_value|
          cell.formatted_value.should == exact_value
        end
      end
    end
  end

  describe '#from_data' do
    let :data do File.read(FixtureFiles::Book::SIMPLE_TEST) end

    it 'identical contents' do
      book_1 = Office::ExcelWorkbook.from_data data
      book_2 = Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST)
      assert_equal book_1.sheets.first.to_csv, book_2.sheets.first.to_csv
    end

    it 'ascii-8bit' do
      data.encoding.should == Encoding::UTF_8
      data.force_encoding Encoding::ASCII_8BIT
      data.encoding.should == Encoding::ASCII_8BIT

      book_1 = Office::ExcelWorkbook.from_data(data)
      book_2 = Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST)
      assert_equal book_1.sheets.first.to_csv, book_2.sheets.first.to_csv
    end

    it 'utf8' do
      data.encoding.should == Encoding::UTF_8
      book_1 = Office::ExcelWorkbook.from_data(data)
      book_2 = Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST)
      assert_equal book_1.sheets.first.to_csv, book_2.sheets.first.to_csv
    end
  end

  describe '#to_data' do
    it 'identical contents' do
      data = Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST).to_data
      assert data.length > 0

      book_1 = Office::ExcelWorkbook.from_data(data)
      book_2 = Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST)
      assert_equal book_1.sheets.first.each_cell.map(&:value), book_2.sheets.first.each_cell.map(&:value)
    end

    it 'ascii-8bit' do
      data = Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST).to_data
      assert data.length > 0
      data.encoding.should == Encoding::ASCII_8BIT

      book_1 = Office::ExcelWorkbook.from_data(data)
      book_2 = Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST)
      assert_equal book_1.sheets.first.each_cell.map(&:value), book_2.sheets.first.each_cell.map(&:value)
    end
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
    book = Office::ExcelWorkbook.new(FixtureFiles::Book::OPENXML_GENERATED)
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

  describe '#clone' do
    let :src do Office::ExcelWorkbook.new(FixtureFiles::Book::LARGE_TEST) end
    let :dst do src.clone end

    it 'objects contain identical xml' do
      src_xmls = src.sheets.map{|sheet| sheet.node.to_xml}
      dst_xmls = dst.sheets.map{|sheet| sheet.node.to_xml}

      # There can be whitespace differences. So to normalise that, pipe them
      # both through nokogiri with noblanks option.
      noblanks = -> xml { Nokogiri::XML(xml, &:noblanks).to_xml }
      srcds = src_xmls.map &noblanks
      dstds = dst_xmls.map &noblanks

      # xml of sheets should now be identical.
      # NOTE not bothering to check other parts'n'stuffs
      dstds.should == srcds
    end

    it 'objects are disparate' do
      dst.object_id.should_not == src.object_id

      objid_of = -> obj, ivar {obj.instance_variable_get(ivar).object_id}

      dst_ivars = dst.instance_variables.map do |ivar_sym|
        objid_of[dst,ivar_sym]
      end

      src_ivars = src.instance_variables.map do |ivar_sym|
        objid_of[src,ivar_sym]
      end

      dst_ivars.should_not == src_ivars
    end
  end
end
