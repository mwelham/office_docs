module Office
  class Logger
    def self.warn(message)
      puts "WARNING: #{message}"
    end
    
    def self.debug_dump(message)
      puts message
    end

    def self.debug_dump_xml(title, xml_doc)
      puts title
      xml_doc.to_xml(:indent => 2).to_s.each_line { |l| puts "  #{l}" }
    end

    def self.debug_dump_table(title, headers, rows, footer = nil)
      column_widths = calc_column_widths(headers, rows)
      total_width = column_widths.inject(column_widths.length + 1) { | width, column_width | width + column_width }
      
      separator = ""
      total_width.times { separator << '-' }

      puts title
      puts "  " + separator
      puts "  " + build_table_row(headers, column_widths)
      puts "  " + separator
      rows.each { |r| puts "  " + build_table_row(r, column_widths)}
      puts "  " + separator
      puts "  " + footer unless footer.nil? or footer.empty?
      puts
    end
    
    private
    
    def self.calc_column_widths(headers, rows)
      column_widths = headers.collect { |h| h.length + 2 }
      rows.each do |row|
        row.each_index do |i|
          if i < column_widths.length
            column_widths[i] = [column_widths[i], row[i].length + 2].max
          else
            column_widths << row[i].length + 2
          end
        end
      end
      column_widths
    end
    
    def self.build_table_row(items, widths)
      row = "|"
      widths.each_index do |i|
        text = i < items.length ? " #{items[i]}" : ""
        row << text
        (widths[i] - text.length).times { row << ' ' }
        row << '|'
      end
      row
    end
  end
end
