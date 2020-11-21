require 'csv'

require 'office/package'
require 'office/constants'
require 'office/errors'
require 'office/logger'

require_relative 'cell'
require_relative 'lazy_cell'
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

    def old_range_to_csv(range: dimension, separator: ',')
      csv = CSV.new '', col_sep: separator, quote_char: ?'
      range.each_rowi do |rowix|
        cell_map = row_cell_nodes_at Location[range.top_left.coli, rowix]
        csv << cell_map[range.top_left.coli..range.bot_rite.coli].each.with_index.map{|cell_node, colix| cell_of(cell_node, colix, rowix).formatted_value}
      end
      csv.string
    end

    def range_to_csv(range: dimension, separator: ',')
      csv = CSV.new '', col_sep: separator, quote_char: ?'
      cell_nodes_of(range, &method(:cell_of)).map do |row_ary|
        csv << row_ary.map(&:formatted_value)
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
        # so just make a unit-sized range
        Office::Range.new "#{dimension_node[:ref]}:#{dimension_node[:ref]}"
      end
    end

    def dimension= range
      dimension_node[:ref] = range.to_s
      @dimension = range
    end

    def calculate_dimension
      # Start with largest and smallest, and contract or expand it to fit the actual cells.
      # Would be nice to start with existing dimension, but it's sometimes not correct.
      # TODO unfortunately nokogiri always instantiates the entire NodeSet, which we don't need.
      # TODO what assumptions hold wrt row and cell ordering? Might be able to use only first and last cells for each row, ¿instead of largest and smallest?
      min, max = data_node.nspath('~row/~c/@r').lazy.reduce [Location.largest, Location.smallest] do |(min,max),r_attr|
        loc = Office::Location.new(r_attr.value)
        # contract and extend min and max, respectively
        [(min & loc), (max | loc)]
      end

      # count empty rows (probably inserted by this codebase)
      min, max = data_node.nspath("~row[count(~c) = 0]/@r").reduce [min,max] do |(min,max),r_attr|
        loc = Office::Location[0, Integer(r_attr.value)]
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
      larger_number_rows = data_node.xpath "xmlns:row[@r >= #{insert_range.top_left.row_r}]"
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
        data_node.add_child(data_node.document.create_element 'row', r: row_r)

        # code to keep the rows in order, which is a more onerous constraint to meet
        #   maybe_existing_row.add_previous_sibling(data_node.document.create_element 'row', r: row_r)
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
      delete_rows = data_node.xpath "xmlns:row[@r >= #{delete_range.top_left.row_r}][@r <= #{delete_range.bot_rite.row_r}]"

      # find all larger rows and decrease
      larger_number_rows = data_node.xpath "xmlns:row[@r > #{delete_range.bot_rite.row_r}]"
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

    # TODO perhaps allow caching of ranges and/or sets, to get the access
    # pattern from higher-level usages.

    # auto-caches row so much faster for things that sequentially access several cells in a row.
    # returns a hash of coli => Cell
    def row_cell_nodes_at loc
      @row_cells ||= Array.new
      @row_cells[loc.rowi] ||= begin
        # Fetch the row node and build cell nodes immediately, otherwise
        # self[loc] usages re-search row cell nodes in a row sequentially from
        # the beginning each time.
        row_node = row_node_at loc

        if row_node
          Hash.new do |h,k|
            h[k] = row_node.element_children[k]
          end
        end
      end
    end

    # Assuming row nodes are never deleted, only moved around.
    #
    # lookup of last row of only 74 rows is around 1ms, cached fetch is around 5-10us
    # so around 100 - 200 times faster.
    def row_node_at loc
      row_node_ix loc.rowi, loc.row_r
    end

    def row_node_ix rowix, row_r = rowix+1
      @row_nodes ||= Array.new
      @row_nodes[rowix] ||= begin
        row_node, = data_node.xpath("xmlns:row[@r=#{row_r}]")
        row_node
      end
    end

    # Pre-fetch a range of cell nodes, as an array of arrays.
    # Call optional blk with (colix, rowix, cell_node), useful for constructing a LazyCell or similar
    def cell_nodes_of range, &blk
      blk ||= -> i,_c,_r {i} # identity if not specified. Slowdown compared to plain value is microseconds at the x10000 scale

      # TODO can be more efficient than this, because each cell requires a hash
      # lookup. Which is fast in ruby, but not as fast as a straightforward
      # iteration of row_node.element_children
      range.each_rowi.map do |rowix|
        range.each_coli.map do |colix|
          blk.call row_node_ix(rowix)&.element_children[colix], colix, rowix
        end
      end
    end

    def lazy_cell_nodes_of range: dimension, &row_blk
      return enum_for :lazy_cell_nodes_of, range: range unless block_given?

      # TODO can be more efficient than this, because each cell requires a hash
      # lookup. Which is fast in ruby, but not as fast as a straightforward
      # iteration of row_node.element_children
      range.each_rowi.each do |rowix|
        # have to construct this with an Enumerator, otherwise 'yield' calls
        # row_blk with the cell. Which is obvs not correct.
        cell_enum = Enumerator.new do |yielder|
          range.each_coli.each do |colix|
            yielder.yield row_node_ix(rowix)&.element_children[colix], colix, rowix
          end
        end
        yield cell_enum
      end
    end

    def cells_of range, &blk
      @cells ||= {}
      lazy_cell_nodes_of range: range do |row_enum|
        row_enum.each do |cell, colix, rowix|
          cell = @cells[[colix, rowix]] ||= cell_of cell_node, colix, rowix
        end
      end
    end

    def strict_cell_of cell_node, colix, rowix
      case [cell_node, colix, rowix]
      in [Nokogiri::XML::Node => cell_node, Integer, Integer]
        Cell.new cell_node, workbook.shared_strings, workbook.styles
      in [nil, Integer => colix, Integer => rowix]
        LazyCell.new self, Location[colix, rowix]
      end
    end

    def loose_cell_of cell_node, colix, rowix
      if cell_node
        Cell.new cell_node, workbook.shared_strings, workbook.styles
      else
        LazyCell.new self, Location[colix, rowix]
      end
    end

    def cell_of cell_node, colix, rowix
      @cells ||= {}
      @cells[[colix, rowix]] ||= loose_cell_of cell_node, colix, rowix
    end

    def invalidate_row_cache
      @cells = {}
      @dimension = nil
      @row_cells = Array.new
      @row_nodes = Array.new
    end

    def [](*args)
      case args
      in [Integer => coli, Integer => rowi]
        self[ Location[coli, rowi] ]

      in [String => a1_location]
        self[ Location.new(a1_location) ]

      in [Location => loc]
        cell_node = row_cell_nodes_at(loc)&.dig(loc.coli)
        cell_of cell_node, *loc

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
      # data_node.children.each do |row_node|
      #   row_node.children.each do |c_node|
      #       yield Cell.new c_node, workbook.shared_strings, workbook.styles
      #   end
      # end

      # comparable to nested each, but slightly cleaner
      # TODO what happens with really huge spreadsheets here?
      data_node.xpath('xmlns:row/xmlns:c').each do |c_node|
        yield cell_of c_node
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

    def each_cache_cell &blk

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
