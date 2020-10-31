require_relative '../lib/office/excel'
require_relative '../lib/office/constants'
require_relative '../lib/office/nokogiri_extensions'

require_relative 'spec_helper'
require_relative 'xml_fixtures'

# copy of the minitest test cases, because specs are easier to zero in on
describe 'ExcelWorkbooksTest' do
  WORKBOOK_PATH = File.join __dir__, '/../test/content/simple-placeholders.xlsx'

  let :simple do
    Office::ExcelWorkbook.new WORKBOOK_PATH
  end

  include XmlFixtures

  it 'replaces one cell' do
    sheet = simple.sheets.first
    cell = sheet.each_cell.first
    cell.value = "This is the new pump"

    Dir.mktmpdir do |dir|
      filename = File.join dir, 'simple.xlsx'
      simple.save filename
      workbook = Office::ExcelWorkbook.new filename
      saved_cell = workbook.sheets.first.each_cell.first
      saved_cell.value.should == cell.value
      saved_cell.node.object_id.should_not == cell.node.object_id
    end
  end

  it 'replaces placeholders' do
    simple = Office::ExcelWorkbook.new WORKBOOK_PATH
    sheet = simple.sheets.first
    sheet.each_cell.lazy.each do |cell|
      # %r otherwise sublime's syntax highlighting breaks
      %r|{{(?<place_name>.*?)}}| =~ cell.value
      next unless place_name
      cell.value = "R-#{place_name.capitalize}"
    end

    Dir.mktmpdir do |dir|
      filename = File.join dir, 'simple.xlsx'
      simple.save filename
      workbook = Office::ExcelWorkbook.new filename
      workbook.sheets.first.each_cell.filter(&:place?).should be_empty
    end
  end

  it 'placeholders in shared strings' do
    vs = {
      horizontal: 'prone',
      flow_rate: 15,
      flow_rate_units: 'm3/sec',
      important: 'Le Grand Fromage',
      broken_place: 'Venterstad',
    }

    plrx = %r|{{(?<place_name>.*?)}}|

    mini_shared_string_doc.nspath('/~sst/~si').each do |si|
      # need Array because Node#search returns a NodeSet
      replace_node, plain_text =
      case si.search('t')
        in []
          # this case may happen if there are 0 t and 0 r children of the si
          node = si.add_child si.document.create_element 't', :'xml:space' => 'preserve'
          [node, '']
        in [node]
          [node, node.text]
        in many
          # we have runs because content model says si can contain exactly 1 t; or multiple r. r can contain exactly 1 t each.
          # so therefore we have many <r><t></t></r> for this case
          # TODO this discards formatting, so we need to 1) unify various formatting attributes somehow, and 2) recreate a r/rPr structure
          # r_nodes = si.search('r')
          # TODO also need to somehow unify xml:space attributes.
          si.children.unlink
          node = si.add_child si.document.create_element 't', :'xml:space' => 'preserve'
          # .text on the many NodeSet will concatenate all text fragments, reconstituting the {{placeholder}}
          [node, many.text]
      end

      replacement = plain_text.gsub plrx do |match| vs[match[2..-3].to_sym] end
      replace_node.children = replacement
    end

    mini_shared_string_doc.nspath('/~sst/~si/~t').text.should == 'Pump InformationproneThis pump moves 15 m3/sec.very Le Grand Fromage thingVenterstad'
  end

  it 'uses Builder to replace element tree' do
    # can also just call methods on bld, but might need to set context first
    bld = Nokogiri::XML::Builder.with replacement_r = dx.create_element('r') do |bld|
      bld.rPr do
        bld.sz val: 10
        bld.rFont val: 'Arial'
        bld.family val: 2
        bld.charset val: 1
      end
      bld.t 'Felix', 'xml:space': 'preserve'
    end

    last_si = mini_shared_string_doc.nspath('/~sst/~si').last
    last_si.children = r_node
    r_ts = mini_shared_string_doc.nspath('~sst/~si/~r/~t')
    r_ts.last.text.should == 'Felix'
    r_ts.size.should == 5
  end

  xit 'pry context' do
    node = mini_shared_string_doc.nspath('/~sst/~si').last

    binding.pry
  end

  it 'multiple xmlns' do
    require 'office/word'
    word = Office::WordDocument.new File.join(__dir__, '../test/content/add_tables_target.docx')
    pdoc = word.get_part "/word/document.xml"
    pdoc.xml.namespaces.should be_any
  end

  let :simple do
    Office::ExcelWorkbook.new File.join(__dir__, '/../test/content/simple-placeholders.xlsx')
  end

  describe Office::Cell do
    describe '#value=' do
      let :sheet do simple.sheets.first end
      let :cell do sheet['A18'] end

      it 'accepts Date' do
        cell.value = 'Hello Darlink'
        cell.node.to_xml.should == "<c r=\"A18\" s=\"0\" t=\"inlineStr\">\n  <is>\n    <t>Hello Darlink</t>\n  </is>\n</c>"
      end

      xit 'accepts boolean'
      xit 'accepts inline String'
      xit 'accepts shared String'
      xit 'accepts Integer'
      xit 'accepts Float'
    end
  end

  describe 'tabular data' do
    let :data do <<~CSV end
      RPM,Discharge PSI,Suction PSI,Net PSI,No,Size,Pitot,GPM,%,Voltage,Amp
      1791,110,10,100,0,Churn,,0,0,"477,476,477","56,55,56"
      1789,106,5,101,3,2.5,7.6,1500,100,"479,475,475","90,86,91"
      1777,96,6,91,3,2.5,17,2250,150,"477,475,474","118,115,118"
      1770,92,5,87,3,8.5,23,3000,170,"452,454,424","128,135,132"
      1762,83,2,82,3,6.1,23,3750,80,"454,424,403","124,128,135"
      1751,65,1,79,3,4.5,23,4500,50,"424,403,389","121,124,128"
    CSV

    let :records do
      require 'csv'
      CSV.parse data
    end

    it 'replaces a range with tabular data' do
      sheet = simple.sheets.first
      doc = sheet.worksheet_part.xml
      # TODO fix nspath to use xmlns if >1 namespace
      # TODO put most of this in Sheet#replace_tabular
      start_cell = sheet['A18']
      start_cell.value.should =~ /tabular/
      range = sheet.merge_ranges.find{|range| range.cover? start_cell.location}
      range.should_not be_nil

      loc_track = start_cell.location.dup
      records.each_with_index do |data_row, rowix|
        # check that range matches data
        data_row.size <= range.width or raise "data too long for #{range}: (#{data_row.size})#{data_row.inspect}"
        data_row.each_with_index do |val, colix|
          loc_track = location = start_cell.location + [colix, rowix]
          # TODO insert second and subsequent rows
          # TODO why does SheetData cache rows?
          # TODO what's the difference between .value = and []= ? The latter would not require LazyCell
          sheet[location].value = Integer(val) rescue val
        end
      end

      # remove merge cells
      sheet.delete_merge_range range

      # TODO update sheet range. Need to find largest cell reference number. Oof.
      # possibly grab current range, extend by loc_track, but then still need to check for rows bumped by insertion
      # unless every row insertion / row deletion / lazy cell insertion tracks the current range.
      # Hmmm.
      sheet.dimension.should == 'A5:K22'
      sheet.calculate_dimension.should == 'A5:K24'
      sheet.dimension = sheet.calculate_dimension
      sheet.dimension.should == 'A5:K24'

      simple.save '/tmp/res.xlsx'
      `localc /tmp/res.xlsx`
    end
  end
end
