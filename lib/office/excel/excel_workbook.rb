module Office
  class ExcelWorkbook < Package
    attr_reader :workbook_part
    attr_reader :shared_strings
    attr_reader :styles
    attr_reader :sheets

    def initialize(filename)
      super(filename)

      @workbook_part = get_relationship_targets(EXCEL_WORKBOOK_TYPE).first
      raise PackageError.new("Excel workbook package '#{@filename}' has no workbook part") if @workbook_part.nil?

      parse_shared_strings
      parse_styles
      parse_workbook_xml
    end

    def self.blank_workbook
      book = ExcelWorkbook.new(File.join(__dir__, '../content/blank.xlsx'))
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

    def save(filename)
      sheets.each(&:update_dimension_node)
      super
    end

    def parse_shared_strings
      case (string_tables = @workbook_part.get_relationship_targets(EXCEL_SHARED_STRINGS_TYPE)).size
      when 0
        binding.pry
      when 1
        @shared_strings = SharedStringTable.new(string_tables.first)
      else
        raise "too many string tables"
      end
    end

    def parse_styles
      styles_part = @workbook_part.get_relationship_targets(EXCEL_STYLES_TYPE).first
      @styles = StyleSheet.new(styles_part) unless styles_part.nil?
    end

    def parse_workbook_xml
      ns_prefix = Package.xpath_ns_prefix(@workbook_part.xml.root)
      @sheets_node = @workbook_part.xml.at_xpath("/#{ns_prefix}:workbook/#{ns_prefix}:sheets")
      raise PackageError.new("Excel workbook '#{@filename}' is missing sheets container") if @sheets_node.nil?

      @sheets = []
      @sheets_node.xpath("#{ns_prefix}:sheet").each { |s| @sheets << Sheet.new(s, self) }
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
      File.open(File.join(__dir__, '../content/empty_sheet.xml')) do |file|
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

      @sheets_node.at_xpath("./#{Package.xpath_ns_prefix(@sheets_node)}:sheet[@name='#{sheet.name}']").remove
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
end
