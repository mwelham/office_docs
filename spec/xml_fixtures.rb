module XmlFixtures
  # I'm sure there's a way to use Rspec let :mini_shared_string_xml in here, if one were inclined that way.

  def mini_shared_string_xml
    @mini_shared_string_xml ||= <<~XML
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst uniqueCount="46" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <!-- text only in string -->
      <si>
        <t xml:space="preserve">Pump Information</t>
      </si>
      <!-- single placeholder only in string -->
      <si>
        <t xml:space="preserve">{{horizontal}}</t>
      </si>
      <!-- multiple placeholder in one string -->
      <si>
        <t xml:space="preserve">This pump moves {{flow_rate}} {{flow_rate_units}}.</t>
      </si>
      <!-- placeholder with formatting -->
      <si>
        <r>
          <rPr>
            <sz val="10"/>
            <rFont val="Arial"/>
            <family val="2"/>
            <charset val="1"/>
          </rPr>
          <t xml:space="preserve">very </t>
        </r>
        <r>
          <rPr>
            <b val="true"/>
            <sz val="10"/>
            <rFont val="Arial"/>
            <family val="2"/>
            <charset val="1"/>
          </rPr>
          <t xml:space="preserve">{{important}}</t>
        </r>
        <r>
          <rPr>
            <sz val="10"/>
            <rFont val="Arial"/>
            <family val="2"/>
            <charset val="1"/>
          </rPr>
          <t xml:space="preserve"> </t>
        </r>
        <r>
          <rPr>
            <i val="true"/>
            <sz val="10"/>
            <rFont val="Arial"/>
            <family val="2"/>
            <charset val="1"/>
          </rPr>
          <t xml:space="preserve">thing</t>
        </r>
      </si>
      <!-- placeholder broken across elements -->
      <si>
        <r>
          <rPr>
            <sz val="10"/>
            <rFont val="Arial"/>
            <family val="2"/>
            <charset val="1"/>
          </rPr>
          <t xml:space="preserve">{{</t>
        </r>
        <r>
          <rPr>
            <b val="true"/>
            <sz val="10"/>
            <rFont val="Arial"/>
            <family val="2"/>
            <charset val="1"/>
          </rPr>
          <t xml:space="preserve">broken</t>
        </r>
        <r>
          <rPr>
            <sz val="10"/>
            <rFont val="Arial"/>
            <family val="2"/>
            <charset val="1"/>
          </rPr>
          <t xml:space="preserve">_place}}</t>
        </r>
      </si>
    </sst>
    XML
  end

  def mini_shared_string_doc
    @mini_shared_string_doc ||= Nokogiri::XML.parse mini_shared_string_xml
  end

  def mini_ts_only
    @mini_ts_only ||= ["Pump Information", "{{horizontal}}", "This pump moves {{flow_rate}} {{flow_rate_units}}."]
  end

  def small_xlsx_sheet_xml
    @small_xlsx_sheet_xml ||= <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheetPr filterMode="false">
          <pageSetUpPr fitToPage="false"/>
        </sheetPr>
        <dimension ref="A5:K24"/>
        <sheetViews>
          <sheetView showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="150" zoomScaleNormal="150" zoomScalePageLayoutView="100" workbookViewId="0">
            <selection pane="topLeft" activeCell="D19" activeCellId="0" sqref="D19"/>
          </sheetView>
        </sheetViews>
        <sheetFormatPr defaultColWidth="11.66015625" defaultRowHeight="12.8" zeroHeight="false" outlineLevelRow="0" outlineLevelCol="0"/>
        <cols>
          <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="1" min="1" style="0" width="12.31"/>
        </cols>
        <sheetData>
          <row r="5" customFormat="false" ht="12.8" hidden="false" customHeight="true" outlineLevel="0" collapsed="false">
            <c r="A5" s="1" t="s">
              <v>0</v>
            </c>
            <c r="B5" s="2" t="s">
              <v>1</v>
            </c>
            <c r="C5" s="2"/>
            <c r="D5" s="0" t="s">
              <v>2</v>
            </c>
            <c r="E5" s="3" t="s">
              <v>3</v>
            </c>
            <c r="F5" s="3"/>
            <c r="G5" s="0" t="s">
              <v>4</v>
            </c>
            <c r="H5" s="0" t="s">
              <v>5</v>
            </c>
          </row>
          <row r="6" customFormat="false" ht="12.8" hidden="false" customHeight="false" outlineLevel="0" collapsed="false">
            <c r="A6" s="1"/>
            <c r="B6" s="0" t="s">
              <v>6</v>
            </c>
            <c r="C6" s="0" t="s">
              <v>7</v>
            </c>
            <c r="D6" s="0" t="s">
              <v>8</v>
            </c>
            <c r="E6" s="0" t="s">
              <v>9</v>
            </c>
            <c r="F6" s="0" t="s">
              <v>10</v>
            </c>
            <c r="G6" s="0" t="s">
              <v>11</v>
            </c>
            <c r="H6" s="0" t="s">
              <v>12</v>
            </c>
          </row>
        </sheetData>
        <mergeCells count="9">
          <mergeCell ref="A5:A9"/>
        </mergeCells>
        <printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>
        <pageMargins left="0.7875" right="0.7875" top="1.05277777777778" bottom="1.05277777777778" header="0.7875" footer="0.7875"/>
        <pageSetup paperSize="9" scale="100" firstPageNumber="1" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" orientation="portrait" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"/>
      </worksheet>
    XML
  end

  def small_xlsx_sheet_doc
    @small_xlsx_sheet_doc ||= Nokogiri::XML.parse small_xlsx_sheet_xml
  end
end
