require 'office/package'
require 'office/constants'
require 'office/errors'
require 'office/logger'

module Office
  class ExcelWorkbook < Package
    attr_accessor :workbook_part
    attr_accessor :sheets
    
    def initialize(filename)
      super(filename)

      @workbook_part = get_relationship_target(EXCEL_WORKBOOK_TYPE)
      raise PackageError.new("Excel workbook package '#{@filename}' has no workbook part") if @workbook_part.nil?
      parse_workbook_xml
    end

    def self.blank_workbook
      ExcelWorkbook.new(File.join(File.dirname(__FILE__), 'content', 'blank.xlsx'))
    end

    def parse_workbook_xml
      @sheets_node = @workbook_part.xml.at_xpath("/xmlns:workbook/xmlns:sheets")
      raise PackageError.new("Excel workbook '#{@filename}' is missing sheets container") if @sheets_node.nil?

      @sheets = []
      @sheets_node.xpath("xmlns:sheet").each { |s| @sheets << Sheet.new(s, self) }
    end

    def debug_dump
      super
      rows = @sheets.collect { |s| ["#{s.name}", "#{s.id}", "#{s.worksheet_part.name}"] }
      Logger.debug_dump_table("Excel Workbook Sheets", ["Name", "Sheet ID", "Part"], rows)
      
      @sheets.each { |s| s.sheet_data.debug_dump }
    end
  end
  
  class Sheet
    attr_accessor :workbook_node
    attr_accessor :name
    attr_accessor :id
    attr_accessor :worksheet_part
    attr_accessor :sheet_data

    def initialize(sheet_node, workbook)
      @workbook_node = sheet_node
      @name = sheet_node["name"]
      @id = sheet_node["sheetId"]
      @worksheet_part = workbook.workbook_part.get_relationship_by_id(sheet_node["id"]).target_part
      
      data_node = worksheet_part.xml.at_xpath("/xmlns:worksheet/xmlns:sheetData")
      raise PackageError.new("Excel worksheet '#{@name} in workbook '#{workbook.filename}' has no sheet data") if data_node.nil?
      @sheet_data = SheetData.new(data_node, self, workbook)
    end
  end
  
  class SheetData
    attr_accessor :node
    attr_accessor :sheet
    attr_accessor :workbook
    attr_accessor :rows
    
    def initialize(node, sheet, workbook)
      @node = node
      @sheet = sheet
      @workbook = workbook

      @rows = []
      node.xpath("xmlns:row").each { |r| @rows << Row.new(r) }
    end
    
    def debug_dump
      Logger.debug_dump_xml("Excel Sheet #{@sheet.worksheet_part.name}", @node)
    end
  end

  class Row
    attr_accessor :node
    attr_accessor :number
    attr_accessor :spans
    attr_accessor :cells
    
    def initialize(row_node)
      @node = row_node
      
      @number = row_node["r"]
      @spans = row_node["spans"]
      
      @cells = []
      node.xpath("xmlns:c").each { |c| @cells << Cell.new(c) }
    end
  end

  class Cell
    attr_accessor :node
    attr_accessor :location
    attr_accessor :style
    attr_accessor :data_type
    attr_accessor :value_node
 
    def initialize(c_node)
      @node = c_node
      @location = c_node["r"]
      @style = c_node["s"]
      @data_type = c_node["t"]
      @value_node = c_node.at_xpath("xmlns:v")
    end
    
    def value
      @value_node.nil? ? nil : @value_node.text
    end
  end
end
