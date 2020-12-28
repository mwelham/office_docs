require_relative 'spec_helper'

require 'yaml'
require_relative '../lib/office/excel/template.rb'

describe Excel::Template do
  let :book do Office::ExcelWorkbook.new BookFiles::SIMPLE_PLACEHOLDERS end
  let :data do
    data = YAML.load_file BookFiles.content_path + 'placeholder-data.yml'
    # convert tabular data to hashy data
    data[:streams] = Excel::Template.tabular_hashify data[:streams]
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
      reload_workbook book do |book|
        `localc #{book.filename}`
      end
    end
  end

  describe described_class::Evaluator do
    let :data do
      data = {controller: {streams: [{start: 'the party'}, {q1: 'steel', "q1:era": 'damascus'}]}}
      data.extend(described_class)
    end

    it 'normal' do
      data.evaluate('controller.streams[0].start').should == 'the party'
    end

    it 'wrong index' do
      data.evaluate('controller.streams[1999].start').should be_nil
    end

    it 'spaces' do
      data.dig(:controller, :streams, 0)[:'the word'] = 'wyrd'
      data.evaluate('controller.streams[0].the word').should == 'wyrd'
    end

    it 'weirdness' do
      data.evaluate('controller..streams[0]].start').should == 'the party'
    end

    it 'colons in names' do
      data.evaluate('controller.streams[1].q1').should == 'steel'
      data.evaluate('controller.streams[1].q1:era').should == 'damascus'
    end

    it 'attributes' do
      data = {first_part: {q1: 't5', :'q1:timestamp' => '2020-12-10 23:01:15'}}
      data.extend(described_class)
      data.evaluate('first_part.q1').should == 't5'
      data.evaluate('first_part.q1:timestamp').should == '2020-12-10 23:01:15'
    end

    it 'single' do
      data.evaluate('c').should be_nil
    end

    it 'raises' do
      ->{data.evaluate('').should == 'the party'}.should raise_error(/invalid expression/i)
    end

    it 'returns nil' do
      expr = 'controller.circuits[0].start'
      data.evaluate(expr).should be_nil
    end
  end

  # uncomment when ExcelWorkbook#dup implemented
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
