require_relative 'spec_helper'

require 'yaml'
require_relative '../lib/office/excel'
require_relative '../lib/office/excel/template.rb'
require_relative '../lib/office/excel/placeholder.rb'

require_relative 'package_debug'

describe Excel::Template do
  using PackageDebug

  let :book do Office::ExcelWorkbook.new FixtureFiles::Book::PLACEHOLDERS end
  let :book2 do Office::ExcelWorkbook.new FixtureFiles::Book::PLACEHOLDERS_TWO end
  let :image do Magick::ImageList.new FixtureFiles::Image::TEST_IMAGE end
  let :sheet do book.sheets.first end

  let :data do
    data = YAML.load_file FixtureFiles::Yaml::PLACEHOLDER_DATA
    # convert tabular data to hashy data
    data[:streams] = Excel::Template.tabular_hashify data[:streams]
    # data.extend Excel::Template.
    # append an image
    # TODO what kind of value will be here from the rails/forms side?
    data[:logo] = image
    data
  end

  describe '.render!' do
    it 'replaces placeholders' do
      placeholders = ->{ book.sheets.flat_map {|sheet| sheet.each_placeholder.to_a } }

      placeholders.call.should_not be_empty
      Excel::Template.render!(book, data)
      placeholders.call.should be_empty
    end

    it 'modifies input book' do
      target_book = Excel::Template.render!(book, data)
      target_book.object_id.should == book.object_id
    end

    it 'displays', display_ui: true do
      Excel::Template.render!(book, data)
      # reload_workbook book2 do |book| `localc #{book.filename}` end
      reload_workbook book do |book| `localc #{book.filename}` end
    end
  end

  describe 'replacement' do
    it 'chooses appropriate data type' do
      cell = book.sheets.first['B11']
      cell.value = "{{important_first_date}}"
      ph = cell.placeholder
      ph[] = Date.today
      cell.to_ruby.should == Date.today
    end

    it 'placeholder does partial replacement' do
      cell = book.sheets.first['B11']
      cell.value.should == "very {{important}} thing"
      ph = cell.placeholder
      ph[] = "we can carry"
      cell.value.should == "very we can carry thing"
    end

    it 'placeholder does partial replacement of non-String' do
      cell = book.sheets.first['B11']
      cell.value.should == "very {{important}} thing"
      ph = cell.placeholder
      date = Date.today
      ph[] = date
      cell.value.should == "very #{date.to_s} thing"
    end

    it 'image insertion replacement' do
      # H20 and L20
      cell = sheet['H20']
      cell.value.should == "{{logo|133x100}}"

      # get image data
      placeholder = Office::Placeholder.parse cell.placeholder.to_s

      image = data.dig *placeholder.field_path

      # add image to sheet
      sheet.add_image image, cell.location, extent: placeholder.image_extent
      cell.value = nil

      # post replacement
      cell.value.should be_nil

      sheet['H20'].value.should be_nil
      book.parts['/xl/drawings/drawing1.xml'].should be_a(Office::XmlPart)
    end
  end

  describe '.distribute' do
    describe 'plain array' do
      let :placeholder_source do
        YAML.load File.read(FixtureFiles::Yaml::PLACEHOLDER_DATA), symbolize_names: true
      end

      it 'refuses to handle array of array values' do
        ->{Excel::Template.distribute fields: placeholder_source[:streams]}.should raise_error(/cannot handle non-hash/)
      end
    end

    describe 'deeply nested' do
      let :data do YAML.load File.read(FixtureFiles::Yaml::MARINE), symbolize_names: true end

      it 'succeeds' do
        rows = []
        Excel::Template.distribute(data) {|row| rows << row }
        rows.size.should == 27
        rows.map(&:size).should == [46, 46, 61, 61, 61, 76, 76, 76, 76, 87, 98, 109, 120, 143, 150, 157, 164, 171, 178, 185, 141, 150, 157, 164, 171, 178, 185]
      end
    end
  end

  describe '.render_tabular' do
    let :placeholder_source do
      YAML.load File.read(FixtureFiles::Yaml::PLACEHOLDER_DATA)
    end

    let :streams_data do
      # convert headers to strings for easier comparisons with data retrieved
      # from sheet.
      headers, *data = placeholder_source[:streams]
      [headers.map(&:to_s), *data]
    end

    let :streams_data_only do placeholder_source[:streams][1..-1] end

    describe 'simple table' do
      it 'overwrite' do
        cell = sheet['A18']
        placeholder = Office::Placeholder.parse cell.placeholder.to_s
        cell.value.should == '{{streams|tabular}}'
        values = data.dig *placeholder.field_path

        old_range = cell.location * [values.first.size, values.size]
        old_cell_data = sheet.cells_of(old_range, &:to_ruby)

        sheet.dimension.to_s.should == 'A5:K26'
        range = Excel::Template.render_tabular sheet, cell, placeholder, values

        sheet.invalidate_row_cache
        # same dimension as previous, so no insertion done
        sheet.dimension.to_s.should == 'A5:K26'

        # fetch the data
        sheet.cells_of(range, &:to_ruby).should == streams_data_only
        sheet.cells_of(range, &:to_ruby).should_not == old_cell_data
      end

      it 'overwrite with headers' do
        placeholder = Office::Placeholder.parse '{{streams|tabular,headers}}'
        values = data.dig *placeholder.field_path

        cell = sheet['A18']
        old_range = cell.location * [values.first.size, values.size]
        old_cell_data = sheet.cells_of(old_range, &:to_ruby)

        sheet.dimension.to_s.should == 'A5:K26'
        range = Excel::Template.render_tabular sheet, cell, placeholder, values

        sheet.invalidate_row_cache
        # same dimension as previous, so no insertion done
        sheet.dimension.to_s.should == 'A5:K26'

        # fetch the data
        sheet.cells_of(range, &:to_ruby).should == streams_data
        sheet.cells_of(range, &:to_ruby).should_not == old_cell_data
      end

      it 'insert' do
        cell = sheet['A18']
        placeholder = Office::Placeholder.parse '{{streams|tabular,insert}}'
        values = data.dig *placeholder.field_path

        # simple data so we can hack ou the size of the insert range fairly easily
        old_range = cell.location * [values.first.size, values.size]
        old_cell_data = sheet.cells_of(old_range, &:to_ruby)

        # this should end up being enlarged
        sheet.dimension.to_s.should == 'A5:K26'

        new_range = Excel::Template.render_tabular sheet, cell, placeholder, values
        new_range.should == old_range

        # validate the data
        sheet.invalidate_row_cache
        sheet.cells_of(new_range, &:to_ruby).should_not == old_cell_data

        # larger than the previous dimension
        sheet.dimension.to_s.should == 'A5:K29'
      end

      it 'vertical overwrite' do
        cell = sheet['A18']
        placeholder = Office::Placeholder.parse '{{streams|tabular,vertical}}'
        values = data.dig *placeholder.field_path

        old_range = cell.location * [values.first.size, values.size]
        old_cell_data = sheet.cells_of(old_range, &:to_ruby)

        sheet.dimension.to_s.should == 'A5:K26'
        new_range = Excel::Template.render_tabular sheet, cell, placeholder, values

        new_range.should == old_range

        sheet.invalidate_row_cache
        sheet.dimension.to_s.should == 'A5:K26'

        # fetch the data
        sheet.cells_of(new_range, &:to_ruby).should == streams_data_only
      end

      it 'default to tabular for non-singular value' do
        placeholder = Office::Placeholder.parse '{{streams}}'
        values = data.dig *placeholder.field_path

        cell = sheet['A18']
        sheet.dimension.to_s.should == 'A5:K26'
        range = Excel::Template.render_tabular sheet, cell, placeholder, values

        sheet.invalidate_row_cache
        # same dimension as previous, so no insertion done
        sheet.dimension.to_s.should == 'A5:K26'

        # fetch the data
        sheet.cells_of(range, &:to_ruby).should == streams_data_only
      end
    end

    describe 'two-branch table' do
      let :pet_data do
        # hack in some higher-level keys
        {fields: {peer_group: {review: YAML.load(File.read(FixtureFiles::Yaml::PETS), symbolize_names: true)}}}
      end

      let :expected do
        YAML.load <<~YAML
        - [organisation,Acme Pet Repairs,Acme Pet Repairs,Acme Pet Repairs,Acme Pet Repairs,Acme Pet Repairs,Acme Pet Repairs,Acme Pet Repairs,Acme Pet Repairs,Acme Pet Repairs,Acme Pet Repairs]
        - [address,1 Seuss Rd,1 Seuss Rd,1 Seuss Rd,1 Seuss Rd,1 Seuss Rd,1 Seuss Rd,1 Seuss Rd,1 Seuss Rd,1 Seuss Rd,1 Seuss Rd]
        - [recorded,2021-05-20,2021-05-20,2021-05-20,2021-05-20,2021-05-20,2021-05-20,2021-05-20,2021-05-20,2021-05-20,2021-05-20]
        - [precise,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00,!ruby/object:DateTime 2021-05-20 07:35:59.000000000 -04:00]
        - [clients.first_name,John,John,Colleen,Colleen,null,null,null,null,null,null]
        - [clients.last_name,Anderson,Anderson,MacKenzie,MacKenzie,null,null,null,null,null,null]
        - [clients.pets.name,Charlie,Feather,Jock,Paddy,null,null,null,null,null,null]
        - [clients.pets.species,cat,cat,dog,dog,null,null,null,null,null,null]
        - [clients.pets.born,2009-05-20,2009-05-20,2019-08-12,2017-10-06,null,null,null,null,null,null]
        - [suppliers.name,null,null,null,null,Hill Scientific Method,Hill Scientific Method,Petrova,Green Industries,Green Industries,Second Life Foods]
        - [suppliers.products.name,null,null,null,null,kibbles, biscuits, collar, catnip, grass, bones]
        YAML
      end

      it 'horizontal overwrite' do
        cell = sheet['B18']

        sheet.dimension.to_s.should == 'A5:K26'

        # post-overwrite marker line
        sheet['A29'].value = 'last'
        sheet['B29'].value = 'line'
        sheet['C29'].value = 'after'
        sheet['K29'].value = 'render'

        placeholder = Office::Placeholder.parse '{{fields.peer_group.review|tabular,horizontal,headers}}'
        values = pet_data.dig *placeholder.field_path

        range = Excel::Template.render_tabular sheet, cell, placeholder, values
        range.to_s.should == "B18:L28"

        sheet.invalidate_row_cache
        sheet.dimension.to_s.should == 'A5:L29'

        # fetch the data
        sheet.cells_of(range, &:to_ruby).should == expected

        # last row of overwritten data
        range.bot_left.to_s.should == 'B28'
        # final row should be unchanged
        sheet.cells_of(Office::Range.new('A29:K29'), &:to_ruby).first.should == ['last', 'line', 'after', nil, nil, nil, nil, nil, nil, nil, 'render']
      end

      it 'horizontal insert' do
        cell = sheet['A18']
        placeholder = Office::Placeholder.parse '{{fields.peer_group.review|tabular,horizontal,headers,insert}}'
        values = pet_data.dig *placeholder.field_path

        sheet.dimension.to_s.should == 'A5:K26'
        range = Excel::Template.render_tabular sheet, cell, placeholder, values
        range.to_s.should == "A18:K28"

        sheet.invalidate_row_cache
        sheet.dimension.to_s.should == "A5:K36"

        # fetch the data
        sheet.cells_of(range, &:to_ruby).should == expected

        # last row of overwritten data
        range.bot_left.to_s.should == 'A28'

        # subsequent row after placeholder should now be first row after insert
        # can't use to_ruby here because the cells have formats in them that we don't understand.
        sheet.cells_of(Office::Range.new('A29:D29'), &:value).first.should == ["1", "44132", "7", "3.141528"]
        sheet.cells_of(Office::Range.new('H30:K30'), &:value).first.should == ['{{logo|133x100}}', nil, nil, '{{logo}}']
      end

      it 'does not insert values whose placeholders have been overwritten by tabular' do
        cell = sheet['A18']

        # there should be some image placeholders
        placeholder_cells = sheet.each_placeholder.to_a
        image_placeholder_cells = placeholder_cells.select{|c| c.placeholder.to_s =~ /logo/}
        image_placeholder_cells.size.should == 2

        cell.value = '{{fields.peer_group.review|tabular,horizontal,headers,overwrite}}'
        Excel::Template.render! book, data.merge(pet_data)

        # no images have been added, because the image placeholders were overwritten
        book.parts['/xl/drawings/drawing1.xml'].should be_nil
      end

      describe 'view' do
        include ReloadWorkbook

        it 'horizontal insert', display_ui: true do
          sheet['B18'].value = '{{fields.peer_group.review|tabular,horizontal,headers,insert}}'
          sheet.accept!(Office::Location.new('A18'), [%w[Tabular-Data T1 T2 T3]].transpose)

          reload_workbook book do |book| `localc #{book.filename}` end

          Excel::Template.render! book, data.merge(pet_data)
          reload_workbook book do |book| `localc #{book.filename}` end

          sheet['A18'].value.should == 'Tabular-Data'
          sheet.cells_of(Office::Range.new('A29:D29')).flatten.map(&:value).should == %w[T1 44132 7 3.141528]
        end
      end
    end
  end

  describe described_class::Evaluator do
    # make data understand evaluate
    def evaluatorize data
      data.extend(described_class)

      # really a testing convenience. Should go eventually.
      def data.split_evaluate expr_str
        # split on . and [] so "streams[0].start" becomes [:streams, 0, :start]
        field_path = expr_str.split(/[\[.\]]+/).map! do |part|
          # convert to integer, then symbol
          Integer part rescue part.to_sym
        end
        self.evaluate field_path
      end

      data
    end

    let :data do
      evaluatorize controller: {streams: [{start: 'the party'}, {q1: 'steel', "q1:era": 'damascus'}]}
    end

    it 'normal' do
      data.split_evaluate('controller.streams[0].start').should == 'the party'
    end

    it 'wrong index' do
      ->{data.split_evaluate('controller.streams[1999].start')}.should raise_error(Excel::Template::PathNotFound, /not found in data/)
    end

    it 'spaces' do
      data.dig(:controller, :streams, 0)[:'the word'] = 'wyrd'
      data.split_evaluate('controller.streams[0].the word').should == 'wyrd'
    end

    it 'weirdness' do
      data.split_evaluate('controller..streams[0]].start').should == 'the party'
    end

    it 'colons in names' do
      data.split_evaluate('controller.streams[1].q1').should == 'steel'
      data.split_evaluate('controller.streams[1].q1:era').should == 'damascus'
    end

    it 'attributes' do
      data = evaluatorize first_part: {q1: 't5', :'q1:timestamp' => '2020-12-10 23:01:15'}

      data.split_evaluate('first_part.q1').should == 't5'
      data.split_evaluate('first_part.q1:timestamp').should == '2020-12-10 23:01:15'
    end

    it 'single nil allowed' do
      data.split_evaluate('c').should be_nil
    end

    it 'last step nil allowed' do
      data.split_evaluate('controller.streams[0].oops').should be_nil
    end

    it 'raises for empty expression' do
      ->{data.split_evaluate('')}.should raise_error(/invalid/i)
    end

    it 'not found' do
      expr = 'controller.circuits[0].start'
      ->{data.split_evaluate(expr)}.should raise_error(Excel::Template::PathNotFound, /not found in data/)
    end
  end

  describe '.render' do
    it 'calls render!' do
      Excel::Template.should_receive :render!
      Excel::Template.render(book, data)
    end

    it 'preserves input book' do
      target_book = Excel::Template.render(book, data)
      target_book.object_id.should_not == book.object_id

      # verify sheet objects are dissimilar - ie their intersection is empty
      (target_book.sheets.map(&:object_id) & book.sheets.map(&:object_id)).should be_empty

      # verify xml parent nodes are dissimilar
      target_book.sheets.map{|sheet| sheet.node.object_id }.should_not == book.sheets.map{|sheet| sheet.node.object_id }

      # in case xml needs to be eyeballed
      # File.write '/tmp/bk.xml', book.sheets.first.node.to_xml
      # File.write '/tmp/tg.xml', target_book.sheets.first.node.to_xml
      # meld <(xmllint --format /tmp/bk.xml) <(xmllint --format /tmp/tg.xml)
    end
  end
end
