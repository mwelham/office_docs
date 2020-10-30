module Office
  class Sheet
    attr_reader :workbook_node
    attr_reader :name
    attr_reader :id
    attr_reader :worksheet_part
    attr_reader :sheet_data

    def initialize(sheet_node, workbook)
      @workbook_node = sheet_node
      @name = sheet_node["name"]
      @id = sheet_node["sheetId"].to_i
      @worksheet_part = workbook.workbook_part.get_relationship_by_id(sheet_node["r:id"]).target_part

      ns_prefix = Package.xpath_ns_prefix(@worksheet_part.xml.root)
      data_node = @worksheet_part.xml.at_xpath("/#{ns_prefix}:worksheet/#{ns_prefix}:sheetData")
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
    attr_reader :node
    attr_reader :sheet
    attr_reader :workbook
    attr_reader :rows

    def initialize(node, sheet, workbook)
      @node = node
      @sheet = sheet
      @workbook = workbook

      @rows = []
      node.xpath("#{Package.xpath_ns_prefix(node)}:row").each { |r| @rows << Row.new(r, workbook.shared_strings, workbook.styles) }
    end

    def add_row(data)
      row_node = Row.create_node(@node.document, @rows.length + 1, data, workbook.shared_strings)
      @node.add_child(row_node)
      @rows << Row.new(row_node, workbook.shared_strings, workbook.styles)
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
end
