module Office
  class SheetData
    attr_reader :node
    attr_reader :sheet
    attr_reader :workbook

    def initialize(node, sheet, workbook)
      @node = node
      @sheet = sheet
      @workbook = workbook
    end

    def invalidate
      @rows = nil
    end

    def rows
      @rows ||= begin
        node.xpath("#{Package.xpath_ns_prefix(node)}:row").map{|r| Row.new(r, workbook.shared_strings, workbook.styles) }
      end
    end

    # add data to xml doc, and to the rows collection
    def add_row(data)
      row_node = Row.create_node(@node.document, rows.length + 1, data, workbook.shared_strings, styles: workbook.styles)
      @node.add_child(row_node)
      rows << Row.new(row_node, workbook.shared_strings, workbook.styles)
    end

    def to_csv(separator)
      data = []
      column_count = 0

      rows.each do |r|
        # pad for missing rows where r.number is discontinguous from last r.number
        data.push([]) until data.length > r.number
        # assign this row
        data[r.number] = r.to_ary
        # update column count
        column_count = [column_count, data[r.number].length].max
      end

      # pad all columns to equal length based on column count
      data.each { |d| d.push("") until d.length == column_count }

      csv = ""
      data.each do |d|
        # quote items containing separator
        items = d.map { |i| i.index(separator).nil? ? i : "'#{i}'" }
        csv << items.join(separator) << "\n"
      end
      csv
    end

    def debug_dump
      data = []
      column_count = 1
      rows.each do |r|
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
