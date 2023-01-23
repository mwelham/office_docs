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
    data[:logo] = image
    data
  end

  describe '.render!' do
    include ReloadWorkbook

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

    it 'placeholder error' do
      sheet['A1'].value = '{{ form_name} }}'
      target_book = Excel::Template.render!(book, data)
      sheet['A1'].value.should == "Placeholder parse error: Unexpected } at 0:15 in '{{ form_name} }}'"
    end

    it 'generated placeholders' do
      # this validates forms-specific default that tickles an obscure bug in
      # inline <is><t> which includes newlines
      book = Office::ExcelWorkbook.blank_workbook
      sheet = book.sheets.first
      data = { one: "One", two:  "Two", third_thing: "Third Thing" }

      data.each{|(k,v)| sheet.add_row [v, "{{#{k}}}"] }

      reload_workbook book, 'default_template.xlsx' do |tbook|
        Excel::Template.render!(tbook, data)
        tbook.save tbook.filename
        range = Office::Range.new 'B1:B3'
        saved_values = tbook.sheets.first.cells_of(range, &:formatted_value).flatten
        saved_values.should == data.values
      end
    end

    it 'invalidates formula caches' do
      range = Office::Range.new('A21:I21')
      sheet.cells_of(range){|c| c.node.nxpath('*:v')}.flatten.count.should == range.width
      Excel::Template.render!(book, data)
      sheet.cells_of(range){|c| c.node.nxpath('*:v')}.flatten.should == []
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
      cell = sheet['H14']
      cell.value.should == "{{logo|133x100}}"

      # get image data
      placeholder = Office::Placeholder.parse cell.placeholder.to_s

      image = data.dig *placeholder.field_path

      # add image to sheet
      sheet.add_image image, cell.location, extent: placeholder.image_extent
      cell.value = nil

      # post replacement
      cell.value.should be_nil

      sheet['H14'].value.should be_nil
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

      describe 'insert and sorting' do
        # this is really a pre-condition to the sorting one
        it 'insert dis-orders rows' do
          cell = sheet['A18']
          placeholder = Office::Placeholder.parse '{{streams|tabular,insert}}'
          values = data.dig *placeholder.field_path
          new_range = Excel::Template.render_tabular sheet, cell, placeholder, values

          cell_refs = sheet.data_node.nxpath('*:row/*:c/@r').map(&:to_s)
          sorted_cell_refs = cell_refs.sort_by{|st| Office::Location.new st}
          sorted_cell_refs.should_not == cell_refs
        end

        # this is really testing sorting
        it 'sort_rows_and_cells works after insert' do
          cell = sheet['A18']
          placeholder = Office::Placeholder.parse '{{streams|tabular,insert}}'
          values = data.dig *placeholder.field_path
          new_range = Excel::Template.render_tabular sheet, cell, placeholder, values

          sheet.sort_rows_and_cells

          cell_refs = sheet.data_node.nxpath('*:row/*:c/@r').map(&:to_s)
          sorted_cell_refs = cell_refs.sort_by{|st| Office::Location.new st}
          sorted_cell_refs.should == cell_refs
        end
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

      describe 'images' do
        include ReloadWorkbook

        let :blobs do YAML.load_file FixtureFiles::Yaml::IMAGE_BLOBS end
        let :imgs do blobs.map{|blob| Magick::Image.from_blob(blob).first} end

        it 'renders images' do
          # extend data.streams with images
          data[:streams].each_with_object imgs.each do |stream_hash, en|
            stream_hash[:image] = en.next
          end

          # how many images before render (actually this is the whole book, but doesn't matter)
          count_images_pre = sheet.drawing_part.xml.nxpath('//*:oneCellAnchor').count

          # render data as usual
          range = Excel::Template.render_tabular sheet, sheet['A18'], (Office::Placeholder.parse '{{streams}}'), data
          sheet.invalidate_row_cache

          # validate image presence
          sheet.drawing_part.xml.nxpath('//*:oneCellAnchor').count.should == count_images_pre + 11
        end
      end
    end

    describe 'two-branch table' do
      let :pet_data do
        # hack in some higher-level keys
        {fields: {peer_group: {review: YAML.load(File.read(FixtureFiles::Yaml::PETS), symbolize_names: true)}}}
      end

      let :expected do
        YAML.load(File.read(FixtureFiles::Yaml::HORIZONTAL_ROWS))
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
        sheet.cells_of(Office::Range.new('A32:H33'),&:to_ruby).should ==
         [["Manual Table", nil, nil, nil, nil, nil, nil, nil],
          ["rpm", "discharge", "suction", "net", "no", "size", "pitot", "gpm"]]
       end

      it 'does not insert values whose placeholders have been overwritten by tabular' do
        # there should be some image placeholders
        placeholder_cells = sheet.each_placeholder.to_a
        image_placeholder_cells = placeholder_cells.select{|c| c.placeholder.to_s =~ /logo/}
        image_placeholder_cells.size.should == 2

        # create a tabular placeholder in a position that will overwrite the
        # image placeholders
        sheet[Office::Location.new('B14')] = '{{fields.peer_group.review|tabular,horizontal,headers,overwrite}}'
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
