require 'office/package'
require 'office/constants'
require 'office/errors'
require 'office/logger'

module Office
  class ExcelWorkbook < Package
    attr_accessor :workbook_part
    attr_accessor :shared_strings
    attr_accessor :sheets
    
    def initialize(filename)
      super(filename)

      @workbook_part = get_relationship_targets(EXCEL_WORKBOOK_TYPE).first
      raise PackageError.new("Excel workbook package '#{@filename}' has no workbook part") if @workbook_part.nil?

      parse_shared_strings
      parse_workbook_xml
    end

    def self.blank_workbook
      book = ExcelWorkbook.new(File.join(File.dirname(__FILE__), 'content', 'blank.xlsx'))
      book.filename = nil
      book
    end

    def self.from_data(data)
      file = Tempfile.new('OfficeExcelWorkbook')
      file.binmode
      file.write(data)
      file.close
      begin
        book = ExcelWorkbook.new(file.path)
        book.filename = nil
        return book
      ensure
        file.delete
      end
    end

    def parse_shared_strings
      shared_strings_part = @workbook_part.get_relationship_targets(EXCEL_SHARED_STRINGS_TYPE).first
      @shared_strings = SharedStringTable.new(shared_strings_part) unless shared_strings_part.nil?
    end
    
    def parse_workbook_xml
      @sheets_node = @workbook_part.xml.at_xpath("/xmlns:workbook/xmlns:sheets")
      raise PackageError.new("Excel workbook '#{@filename}' is missing sheets container") if @sheets_node.nil?

      @sheets = []
      @sheets_node.xpath("xmlns:sheet").each { |s| @sheets << Sheet.new(s, self) }
    end

    def add_sheet(name)
      raise PackageError.new("New sheet name cannot be empty") if name.nil? or name.empty?
      sheet_id = 1
      @sheets.each do |s|
        raise PackageError.new("Spreadsheet already contains a sheet named '#{name}'") if name == s.name
        matches = s.worksheet_part.name.scan(/.*\/sheet(\d+)\.xml\z/i)
        sheet_id = [sheet_id, s.id + 1, (matches.nil? || matches.empty? ? 0 : matches[0][0].to_i + 1)].max
      end

      sheet_part = nil
      File.open(File.join(File.dirname(__FILE__), 'content', 'empty_sheet.xml')) do |file|
        sheet_part = add_part("/xl/worksheets/sheet#{sheet_id}.xml", file, XLSX_SHEET_CONTENT_TYPE)
      end
      relationship_id = @workbook_part.add_relationship(sheet_part, EXCEL_WORKSHEET_TYPE)

      node = Sheet.add_node(@sheets_node, name, sheet_id, relationship_id)
      @sheets << Sheet.new(node, self)
      @sheets.last
    end

    def find_sheet_by_name(name)
      @sheets.each { |s| return s if s.name == name }
      nil
    end

    def remove_sheet(sheet)
      return if sheet.nil?
      raise PackageError.new("sheet not found in workbook") unless @sheets.include? sheet

      @sheets_node.at_xpath("./xmlns:sheet[@name='#{sheet.name}']").remove
      remove_part(sheet.worksheet_part)
      @sheets.delete(sheet)
    end

    def debug_dump
      super
      @shared_strings.debug_dump unless @shared_strings.nil?

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
      @id = sheet_node["sheetId"].to_i
      @worksheet_part = workbook.workbook_part.get_relationship_by_id(sheet_node["r:id"]).target_part
      
      data_node = @worksheet_part.xml.at_xpath("/xmlns:worksheet/xmlns:sheetData")
      raise PackageError.new("Excel worksheet '#{@name} in workbook '#{workbook.filename}' has no sheet data") if data_node.nil?
      @sheet_data = SheetData.new(data_node, self, workbook)
    end

    def add_row(data)
      @sheet_data.add_row(data)
    end
    
    def to_csv(separator = ',')
      @sheet_data.to_csv(separator)
    end
    
    def self.add_node(parent_node, name, sheet_id, relationship_id)
      sheet_node = parent_node.document.create_element("sheet")
      parent_node.add_child(sheet_node)
      sheet_node["name"] = name
      sheet_node["sheetId"] = sheet_id.to_s
      sheet_node["r:id"] = relationship_id
      sheet_node
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
      node.xpath("xmlns:row").each { |r| @rows << Row.new(r, workbook.shared_strings) }
    end

    def add_row(data)
      row_node = Row.create_node(@node.document, @rows.length + 1, data, workbook.shared_strings)
      @node.add_child(row_node)
      @rows << Row.new(row_node, workbook.shared_strings)
    end

    def to_csv(separator)
      data = []
      column_count = 0
      @rows.each do |r|
        data.push([]) until data.length > r.number
        data[r.number] = r.to_ary
        column_count = [column_count, data[r.number].length].max
      end
      data.each { |d| d.push("") until d.length == column_count }

      csv = ""
      data.each do |d|
        items = d.map { |i| i.index(separator).nil? ? i : "'#{i}'" }
        csv << items.join(separator) << "\n"
      end
      csv
    end

    def debug_dump
      data = []
      column_count = 1
      @rows.each do |r|
        data.push([]) until data.length > r.number
        data[r.number] = r.to_ary.insert(0, (r.number + 1).to_s)
        column_count = [column_count, data[r.number].length].max
      end
      
      headers = [ "" ]
      0.upto(column_count - 2) { |i| headers << Cell.column_name(i) }
      
      Logger.debug_dump_table("Excel Sheet #{@sheet.worksheet_part.name}", headers, data)
    end
  end

  class Row
    attr_accessor :node
    attr_accessor :number
    attr_accessor :spans
    attr_accessor :cells
    
    def initialize(row_node, string_table)
      @node = row_node
      
      @number = row_node["r"].to_i - 1
      @spans = row_node["spans"]
      
      @cells = []
      node.xpath("xmlns:c").each { |c| @cells << Cell.new(c, string_table) }
    end

    def self.create_node(document, number, data, string_table)
      row_node = document.create_element("row")
      row_node["r"] = number.to_s unless number.nil?
      
      unless data.nil? or data.length == 0
        row_node["spans"] = "1:#{data.length}"
        0.upto(data.length - 1) do |i| 
          c_node = Cell.create_node(document, number, i, data[i], string_table)
          row_node.add_child(c_node)
        end
      end

      row_node
    end

    def to_ary
      ary = []
      @cells.each do |c|
        ary.push("") until ary.length > c.column_num
        ary[c.column_num] = c.value
      end
      ary 
    end
  end

  class Cell
    attr_accessor :node
    attr_accessor :location
    attr_accessor :style
    attr_accessor :data_type
    attr_accessor :value_node
    attr_accessor :shared_string
 
    def initialize(c_node, string_table)
      @node = c_node
      @location = c_node["r"]
      @style = c_node["s"]
      @data_type = c_node["t"]

      # Originally did this, but was incredibly slow:
      #@value_node = c_node.at_xpath("xmlns:v")
      @value_node = nil
      c_node.elements.each { |e| @value_node = e if e.name == 'v' }

      if is_string? && !@value_node.nil?
        string_id = @value_node.content.to_i
        @shared_string = string_table.get_string_by_id(string_id)
        raise PackageError.new("Excel cell #{@location} refers to invalid shared string #{string_id}") if @shared_string.nil?
        @shared_string.add_cell(self)
      end
    end
    
    def self.create_node(document, row_number, index, value, string_table)
      cell_node = document.create_element("c")
      cell_node["r"] = "#{column_name(index)}#{row_number}"
      
      value_node = document.create_element("v")
      cell_node.add_child(value_node)

      unless value.nil? or value.to_s.empty?
        if value.is_a? Numeric
          value_node.content = value
        else
          cell_node["t"] = "s"
          value_node.content = string_table.id_for_text(value.to_s)
        end
      end
      
      cell_node
    end
    
    
    def is_string?
      data_type == "s"
    end

    def self.column_name(index)
      name = ""
      while index >= 0
        name << ('A'.ord + (index % 26)).chr
        index = index/26 - 1
      end
      name.reverse
    end

    def column_num
      letters = /([a-z]+)\d+/i.match(@location)[1].downcase.reverse

      num = letters[0].ord - 'a'.ord
      1.upto(letters.length - 1) { |i| num += (letters[i].ord - 'a'.ord + 1) * (26 ** i) }
      num
    end
    
    def row_num
      /[a-z]+(\d+)/i.match(@location)[1].to_i - 1
    end
    
    def value
      return nil if @value_node.nil?
      is_string? ? @shared_string.text : @value_node.content
    end
  end
  
  class SharedStringTable
    attr_accessor :node
    
    def initialize(part)
      @node = part.xml.at_xpath("/xmlns:sst")
      # TODO Keep these up-to-date
      @count_attr = @node.attribute("count")
      @unique_count_attr = @node.attribute("uniqueCount")

      @strings_by_id = {}
      @strings_by_text = {}
      node.xpath("xmlns:si").each { |si| parse_si_node(si) }
    end

    def parse_si_node(si)
      string = SharedString.new(si, @strings_by_id.length)
      @strings_by_id[string.id] = string
      @strings_by_text[string.text] = string
      string.id
    end
    
    def get_string_by_id(id)
      @strings_by_id[id]
    end

    def id_for_text(text)
      return @strings_by_text[text].id if @strings_by_text.has_key? text
      
      si = node.document.create_element("si")
      t = node.document.create_element("t")
      t.content = text
      si.add_child(t)
      @node.add_child(si)
      
      parse_si_node(si)
    end
    
    def debug_dump
      rows = @strings_by_id.values.collect do |s|
        cells = s.cells.collect { |c| c.location }
        ["#{s.id}", "#{s.text}", "#{cells.join(', ')}"]
      end
      
      footer = ","
      footer << "  count = #{@count_attr.value}" unless @count_attr.nil?
      footer << "  unique count = #{@unique_count_attr.value}" unless @unique_count_attr.nil?
      Logger.debug_dump_table("Excel Workbook Shared Strings", ["ID", "Text", "Cells"], rows, footer)
    end
  end
  
  class SharedString
    attr_accessor :node
    attr_accessor :text_node
    attr_accessor :id
    attr_accessor :cells

    def initialize(si_node, id)
      @node = si_node
      @id = id
      @text_node = si_node.at_xpath("xmlns:t")
      @cells = []
    end
    
    def text
      text_node.content
    end
    
    def add_cell(cell)
      @cells << cell
    end
  end
end
