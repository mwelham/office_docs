require 'csv'

require 'office/package'
require 'office/constants'
require 'office/errors'
require 'office/logger'

require_relative 'cell'
require_relative 'location'
require_relative 'range'
require_relative 'excel_workbook'

module Office
  class Sheet
    attr_reader :workbook_node
    attr_reader :name
    attr_reader :id
    attr_reader :worksheet_part
    attr_reader :workbook

    def initialize(sheet_node, workbook)
      @workbook = workbook
      @workbook_node = sheet_node
      @name = sheet_node['name']
      @id = Integer sheet_node['sheetId']
      @worksheet_part = workbook.workbook_part.get_relationship_by_id(sheet_node["r:id"]).target_part
    end

    def data_node
      @data_node ||= begin
        ns_prefix = Package.xpath_ns_prefix(worksheet_part.xml.root)
        node = worksheet_part.xml.at_xpath("/#{ns_prefix}:worksheet/#{ns_prefix}:sheetData")
        raise PackageError, "Excel worksheet '#{@name} in workbook '#{workbook.filename}' has no sheet data" if node.nil?
        node
      end
    end

    def sheet_data
      @sheet_data ||= SheetData.new(data_node, self, workbook)
    end

    def add_row(data)
      sheet_data.add_row(data)
    end

    # Excel-compatible csv, which is a little different to standard csv.
    def to_excel_csv(separator = ',')
      sheet_data.to_csv(separator)
    end

    def range_to_csv(range: dimension, separator: ',')
      csv = CSV.new '', col_sep: separator, quote_char: ?'
      colix = range.top_left.coli
      range.each_rowi do |rowix|
        cell_map = row_at Location[colix, rowix]
        csv << cell_map.values.map(&:formatted_value)
      end
      csv.string
    end

    # alias to_csv to_excel_csv
    def to_csv(separator = ?,)
      range_to_csv range: dimension, separator: separator
    end

    # TODO what does this do, exactly? Yes OK, it adds a node. But what node? To where? For what purpose?
    def self.add_node(parent_node, name, sheet_id, relationship_id)
      sheet_node = parent_node.document.create_element("sheet")
      parent_node.add_child(sheet_node)
      sheet_node["name"] = name
      sheet_node["sheetId"] = sheet_id.to_s
      sheet_node["r:id"] = relationship_id
      sheet_node
    end

    def dimension_node
      # this is about 1.2 - 3 times faster, but optimising this is not worthwhile here.
      # node.children.first.children.find{|n| n.name == 'dimension'}
      node.xpath('xmlns:worksheet/xmlns:dimension').first
    end

    def dimension
      # TODO /:/ =~ is nearly as fast
      @dimension ||=
      if dimension_node[:ref].include?(?:)
        # TODO there must be a better way to handle this
        Office::Range.new dimension_node[:ref]
      else
        # sometimes for blank worksheets, dimension_node[:ref] == 'A1'
        calculate_dimension
      end
    end

    def dimension= range
      dimension_node[:ref] = range.to_s
      @dimension = nil
    end

    def calculate_dimension
      # Start with largest and smallest, and contract or expand it to fit the actual cells.
      # Would be nice to start with existing dimension, but it's sometimes not correct.
      # TODO unfortunately nokogiri always instantiates the entire NodeSet, which we don't need.
      # TODO what assumptions hold wrt row and cell ordering? Might be able to use only first and last cells for each row, ¿instead of largest and smallest?
      min, max = sheet_data.node.xpath('xmlns:row/xmlns:c/@r').lazy.inject [Location.largest, Location.smallest] do |(min,max),r_attr|
        loc = Office::Location.new(r_attr.text)
        # contract and extend min and max, respectively
        [(min & loc), (max | loc)]
      end
      Office::Range.new min, max
    end

    # create a Office::Range from Location (or maybe later string A1 and string A1:Z26)
    # convert a Location to a Range
    # TODO maybe also handle strings A1 and A1:Z26
    def to_range obj
      case obj
      when Office::Location
        Office::Range.new obj, obj
      when Office::Range
        obj
      else
        raise "wut!? do what with what rows where!? #{insert_here.inspect}"
      end
    end

    # Insert new rows just before the specified location/range.
    #
    # arg is_a Office::Range (several rows, or one row), or is_a Office::Location (insert 1 row)
    #
    # New row nodes will be returned. As an Array, not as a NodeSet.
    #
    # New row nodes will have only the r= attribute.
    #
    # TODO return actual row nodes inserted, range of locations inserted, or both?
    def insert_rows insert_here
      insert_range = to_range insert_here

      # uh-oh now we have to shuffle all larger rows up by insert_range.height and adjust their cell refs to match
      #
      # NOTE we can't assume that row nodes are in the same order as their r=
      # attributes, and we can't assume that rows have contiguous r= attributes
      # - mainly because other parts of this code don't bother to maintain that
      # constraint.
      #
      # TODO this will break formulas and ranges referring to these cells
      larger_number_rows = sheet_data.node.xpath "xmlns:row[@r >= #{insert_range.top_left.row_r}]"
      larger_number_rows.each do |row_node|
        # increase r for the row_node by the insert_range height
        row_number = Integer(row_node[:r]) + insert_range.height
        row_node[:r] = row_number

        # Correspondingly increase r for each of the cells in the row.
        # Iterating through direct children is faster than using xpath.
        row_node.children.each do |cell|
          next unless cell.name == ?c
          # No need to recalculate the row number here, because we already know it.
          # But we need to keep the column index.
          colst, _rowst = Location.parse_a1 cell[:r]
          cell[:r] = Location.of_r colst, row_number
        end
      end

      # tell sheet data to recalculate next time
      sheet_data.invalidate
      invalidate_row_cache

      # Create new row nodes and return them
      insert_range.each_row_r.map do |row_r|
        # Yes, this actually works, because <row r=val> val can be
        # out-of-order in the xml and localc will handle that.
        #
        # Presumably Excel will also be fine with it..?
        #
        # returns new node added
        sheet_data.node.add_child(sheet_data.node.document.create_element 'row', r: row_r)

        # code to keep the rows in order, which is a more onerous constraint to meet
        #   maybe_existing_row.add_previous_sibling(sheet_data.node.document.create_element 'row', r: row_r)
      end
    end

    # delete the specified set of rows
    #
    # delete_these is either a Location or a Office::Range
    #
    # TODO this is substantially the same as insert rows, except for operators and comparators
    #
    # TODO this does two iterations through the entire sheetData/row nodes.
    # Could reduce that with caching.
    def delete_rows delete_these
      delete_range = to_range delete_these

      # find all rows to be deleted. Range is inclusive.
      delete_rows = sheet_data.node.xpath "xmlns:row[@r >= #{delete_range.top_left.row_r}][@r <= #{delete_range.bot_rite.row_r}]"

      # find all larger rows and decrease
      larger_number_rows = sheet_data.node.xpath "xmlns:row[@r > #{delete_range.bot_rite.row_r}]"
      larger_number_rows.each do |row_node|
        # increase r for the row_node by the delete_range height
        row_number = Integer(row_node[:r]) - delete_range.height
        row_node[:r] = row_number

        # Correspondingly increase r for each of the cells in the row.
        # Iterating through direct children is faster than using xpath.
        row_node.children.each do |cell|
          next unless cell.name == ?c
          # No need to recalculate the row number here, because we already know it.
          # But we need to keep the column index.
          colst, _rowst = Location.parse_a1 cell[:r]
          cell[:r] = Location.of_r colst, row_number
        end
      end

      # tell sheet data to recalculate next time
      sheet_data.invalidate
      invalidate_row_cache

      # TODO remove/update mergeCells referring to these deleted rows
      # TODO update formulas referring to rows moved

      # remove the row nodes
      delete_rows.unlink
    end

    # possibly this family of merge_cells calls should have some kind of wrapper class
    def merge_cells
      worksheet_part.xml.nspath 'xmlns:worksheet/xmlns:mergeCells'
    end

    def merge_ranges
      merge_cells.children.map{|n| Office::Range.new n[:ref]}
    end

    def delete_merge_range range
      # find node and delete it
      merge_cells_node, = node.nspath 'xmlns:worksheet/xmlns:mergeCells'
      to_delete = merge_cells_node.children.find{|n| n[:ref] == range.to_s}
      to_delete or raise "no range found matching #{range}"
      to_delete.unlink
      merge_cells_node[:count] = merge_cells_node.children.count
      to_delete
    end

    # auto-caches row so much faster for things that sequentially access several cells in a row.
    def row_at loc
      @row_cache ||= Array.new
      @row_cache[loc.rowi] ||= begin
        # Fetch the row node and build cell nodes immediately, otherwise
        # self[loc] usages re-search row cell nodes in a row sequentially from
        # the beginning each time.
        row_node, = data_node.xpath("xmlns:row[@r=#{loc.row_r}]")

        if row_node
          # NOTE we assume that cells have r= attributes that are in order and contiguous
          row_node.element_children.map do |cell_node|
            cell = Cell.new cell_node, workbook.shared_strings, workbook.styles
            [cell.location.coli, cell]
          end.to_h
        end
      end
    end

    def invalidate_row_cache
      @dimension = nil
      @row_cache = Array.new
    end

    def [](*args)
      case args
      in [Integer => coli, Integer => rowi]
        self[ Location[coli, rowi] ]

      in [String => a1_location]
        self[ Location.new(a1_location) ]

      in [Location => loc]
        cell_node = row_at(loc)&.dig(loc.coli)
        cell_node || LazyCell.new(self, loc)

      else
        raise "don't know how to get cell from #{location.inspect}"
      end
    end

    # The fastest way to provide all actual cells.
    #
    # Intended for finding placeholders. If you want to do this by row/col,
    # consider Location.new and Sheet#[]
    #
    # This does not guarantee that they will be in row/col or col/row order.
    # Does not guarantee that cells or rows will be contiguous, even if they are in order.
    def each_cell_by_node &blk
      # TODO change __method__ to :each_cell once testing settles down
      return enum_for __method__ unless block_given?
      # sheet_data.node.children.each do |row_node|
      #   row_node.children.each do |c_node|
      #       yield Cell.new c_node, workbook.shared_strings, workbook.styles
      #   end
      # end

      # comparable to nested each, but slightly cleaner
      # TODO what happens with really huge spreadsheets here?
      sheet_data.node.xpath('xmlns:row/xmlns:c').each do |c_node|
        yield Cell.new c_node, workbook.shared_strings, workbook.styles
      end
    end

    def sub range
      # TODO implement this to fetch a nodeset of rows each with cell filtering
      # and otherwise all the sheet machinery. So possibly just generalise Sheet to handle
      # a nodeset instead of a single node
    end

    # iterates by sheet_data.rows : Array<Row>
    def each_row_cell by = :row, &blk
      return enum_for __method__, by unless block_given?

      case by
      when :row
        sheet_data.rows.each do |row| row.cells.each &blk end
      when :col
        raise NotImplementedError, 'iterating cells by column not supported yet'
      end
    end

    alias each_cell each_row_cell
    # alias each_cell each_cell_by_node

    # Create a separate method for this, because there may be a more optimal way
    # of finding placeholders.
    def each_placeholder &blk
      return enum_for __method__ unless block_given?

      each_cell do |cell|
        yield cell if cell.placeholder
      end
    end

    def node; worksheet_part.xml end
    def to_xml; node.to_xml end

    class NlInspect
      def inspect; "\n" end
    end

    # very rough debug display that ignores ranges and things
    def spla
      max_row = sheet_data.rows.map{|r| r.cells.size}.max
      rows = sheet_data.rows.map{|r| cells = r.cells.dup; cells[max_row-1] ||= nil; cells}
      # eaurgh
      row_it = rows.each
      sep_it = [NlInspect.new].cycle
      sep_rows = []
      loop do
        sep_rows << row_it.next
        # sep_rows << sep_it.next
      end
      sep_rows
    end

    # doesn't work properly but not fighting with it now
    def inspect; spla.inspect end

    def []=(coli,rowi,value)
      case sheet_data.rows
      when NilClass
        LazyCell.new(self, coli, rowi).value = value
      else
        rowi.cells[coli].value = value
      end
    end
  end

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

    # TODO redundant for now, but maybe worth having a Bulk Inserter or something like that.
    class Collector
      def initialize
        @collection = []
      end

      def << values
        @collection << values
      end

      def to_a
        @collection
      end
    end

    # blk can be called, possibly several times, with an array of values to insert
    # TODO should values be cells?
    # location can be one of: cell; location string (B32 etc); (x,y) indices
    def replace_tabular location, &blk
      top_left_cell = self[location]
      collector = Collector.new
      yield collector
      latest_row = collector.to_a.last
      added_range = location .. location + {row: latest_row.length}
      if range(added_range).any?
        puts "warning: overwriting values"
      end

      add_row latest_row
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
