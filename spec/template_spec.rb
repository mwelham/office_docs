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

    it 'displays replacements', display_ui: true do
      Excel::Template.render!(book, data)
      # reload_workbook book2 do |book| `localc #{book.filename}` end
      reload_workbook book do |book| `localc #{book.filename}` end
    end
  end

  describe 'replacement' do
    it 'placeholder does partial replacement' do
      cell = book.sheets.first['B11']
      cell.value.should == "very {{important}} thing"
      ph = cell.placeholder
      ph[] = "we can carry"
      cell.value.should == "very we can carry thing"
    end

    it 'image insertion replacement' do
      # H20 and L20
      sheet = book.sheets.first
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

    it 'table no insertion' do
      sheet = book.sheets.first
      cell = sheet['A18']
      cell.value.should == "{{streams|tabular}}"
      placeholder = Office::Placeholder.parse cell.placeholder.to_s

      values = data.dig *placeholder.field_path
      tabular_data = Excel::Template.table_of_hash(values)

      # write data to sheet
      sheet.accept!(cell.location, tabular_data)

      # fetch the data
      sheet.invalidate_row_cache
      range = cell.location * [tabular_data.first.count, tabular_data.count]
      sheet.cells_of(range, &:to_ruby).should == tabular_data
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
