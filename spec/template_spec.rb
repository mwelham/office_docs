require 'yaml'
require 'hash_wrap'

require_relative 'spec_helper'
require_relative '../lib/office/excel/template.rb'

describe Excel::Template do
  let :book do Office::ExcelWorkbook.new BookFiles::SIMPLE_PLACEHOLDERS end
  let :data do
    data = YAML.load_file File.realpath File.join __dir__, '../test/content/placeholder-data.yml'
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
      data = {controller: {streams: [{start: 'the party'}]}}
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
  xdescribe '.render' do
    it 'calls render!' do
      Excel::Template.should_receive :render!
      Excel::Template.render(book, data)
    end

    it 'preserves input book' do
      target_book = Excel::Template.render(book, data)
      target_book.object_id.should_not == book.object_id
    end
  end
end
