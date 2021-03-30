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

    # Make an exact copy completely separated from self. That is, changes to the
    # copy will not affect self, and changes to self will not affect the copy.
    def clone
      Dir.mktmpdir do |dir|
        file_path = (Pathname.new(dir) + 'tmp.xlsx')
        save(file_path)

        Office::ExcelWorkbook.new(file_path).tap do |book|
          book.filename = nil
        end
      end
    end

    def self.from_data(data)
      Dir.mktmpdir do |dir|
        file_path = (Pathname.new(dir) + 'tmp.xlsx')
        file_path.open 'wb:ASCII-8BIT' do |io|
          io.write(data)
        end

        Office::ExcelWorkbook.new(file_path).tap do |book|
          book.filename = nil
        end
      end
    end

    def save(filename)
      sheets.each(&:update_dimension_node)
      super
    end

    def parse_shared_strings
      case (string_tables = @workbook_part.get_relationship_targets(EXCEL_SHARED_STRINGS_TYPE)).size
      when 1
        @shared_strings = SharedStringTable.new(string_tables.first)
      when 0
        # hopefully other things will be fine with this
      else
        raise PackageError, 'too many string tables'
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

    # TODO almost identical to Package#add_image_part_rel
    def add_drawing_part_rel(drawing, part)
      drawing_part = add_drawing_part drawing, part.path_components
      relationship_id = add_relationship part, drawing_part, DRAWING_RELATIONSHIP_TYPE

      [relationship_id, drawing_part]
    end

    # drawing is anything that has a to_xml (that produces drawing-compatible xml)
    # drawing will be added as a part underneath /<path_components>/media/drawingX.xml
    # returns an XmlPart instance
    # TODO almost identical to Package#add_image_part
    def add_drawing_part(drawing, path_components)
      prefix = File.join ?/, path_components, 'drawings/drawing'

      # unused_part_identifier is 1..n : Integer
      part_name = "#{prefix}#{unused_part_identifier prefix}.xml"

      # TODO there must be a way to get an IO from nokogiri
      # TODO massive round-trip - this encodes the nokogiri document as a String, and then XmlPart parses it.
      add_part part_name, StringIO.new(drawing.to_xml), drawing.mime_type
    end
  end
end
