require_relative '../lib/office/nokogiri_extensions'

require_relative 'spec_helper'
require_relative 'xml_fixtures'

# copy of the minitest test cases, because specs are easier to zero in on
describe Nokogiri::XML::Searchable do
  include XmlFixtures

  describe '#nspath' do
    it 'rooted xpath on doc' do
      t = mini_shared_string_doc.nspath '/~sst/~si/~t'
      t.size.should == 3
      t.map(&:text).should == mini_ts_only
    end

    it 'relative xpath on doc' do
      t = mini_shared_string_doc.nspath '~sst/~si/~t'
      t.size.should == 3
      t.map(&:text).should == mini_ts_only
    end

    it 'relative xpath for node' do
      sst = mini_shared_string_doc.nspath('~sst').first
      t = sst.nspath '~si/~t'
      t.size.should == 3
      t.map(&:text).should == mini_ts_only
    end

    # note that atttributes do not require namespaces
    it 'does not interfere with attributes' do
      nodeset = mini_shared_string_doc.nspath '//~rPr/~rFont[@val="Arial"]/@val'
      nodeset.text.should == 'Arial' * 7
    end

    describe 'non-root node with namespace declarations' do
      it 'to be continued'
    end

    describe 'multiple namespaces' do
      # not sure what this means, maybe it will come back
      it 'Nokogiri::Searchable#nspath works for Node' do
        sheet_data_node = small_xlsx_sheet_doc.nspath '/~worksheet/~sheetData'
        nodeset = sheet_data_node.nspath '~row/~c/*'
        nodeset.text.should == '0123456789101112'
      end

      # defaults to xmlns: if there are several namespaces in a doc
      it 'defaults to xmlns: when doc has several namespaces' do
        # xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        # xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        # nodeset = dx.nspath '/xmlns:worksheet/xmlns:sheetData/xmlns:row/xmlns:c'
        nodeset = small_xlsx_sheet_doc.nspath '/~worksheet/~sheetData/~row/~c/*'
        nodeset.text.should == '0123456789101112'
      end
    end

    describe 'no namespace declaration' do
      let :no_namespace_xml do
        <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet>
          <sheetPr filterMode="false">
            <pageSetUpPr fitToPage="false"/>
          </sheetPr>
          <dimension ref="A5:K24"/>
          <sheetViews>
            <sheetView showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="150" zoomScaleNormal="150" zoomScalePageLayoutView="100" workbookViewId="0">
              <selection pane="topLeft" activeCell="D19" activeCellId="0" sqref="D19"/>
            </sheetView>
          </sheetViews>
        </worksheet>
        XML
      end

      it "replaces ~ with ''" do
        dx = Nokogiri::XML.parse no_namespace_xml
        node = dx.nspath '/~worksheet/~sheetViews/~sheetView[position()=1]/~selection[@pane="topLeft"]/@pane'
        node.text.should == 'topLeft'
      end

      it "accepts plain tags" do
        dx = Nokogiri::XML.parse no_namespace_xml
        node = dx.nspath '/worksheet/sheetViews/sheetView[position()=1]/selection/@pane'
        node.text.should == 'topLeft'
      end
    end
  end
end
