require 'csv'
require 'base64'

require 'office/package'
require 'office/constants'
require 'office/errors'
require 'office/logger'
require 'office/nokogiri_extensions'

require_relative 'cell'
require_relative 'lazy_cell'
require_relative 'location'
require_relative 'range'
require_relative 'excel_workbook'
require_relative 'sheet_data'
require_relative 'image_drawing'

module Office
  # The design rationale behind this class is that a Sheet is a *grid* of cells.
  #
  # It therefore focuses on accessing, manipulating, creating cells as addressed
  # by A1 style locations and A1:CC28 style ranges.
  #
  # This naturally creates some tension with the underlying xlsx xml model,
  # which consists of <row> elements containing many cell <c> elements.
  #
  # xpath queries to access the xml nodes, while not slow, are slower than
  # lookups in ruby Array and Hash structures. So whenever possible, lookup of
  # cell <c> nodes are cached in Array and/or Hash structures. Obviously those
  # caches must be invalidated when row lookup indices are changed by
  # row-oriented insert and delete operations which must renumber subsequent row
  # elements.
  class Sheet
    attr_reader :name
    attr_reader :id
    attr_reader :worksheet_part
    attr_reader :workbook

    def initialize(sheet_node, workbook)
      @workbook = workbook
      @name = sheet_node['name']
      @id = Integer sheet_node['sheetId']
      @worksheet_part = workbook.workbook_part.get_relationship_by_id(sheet_node["r:id"]).target_part

      @dimension_fn = method :dimension_of_xlsx
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

    # Backwards compatibility
    def add_row(data)
      sheet_data.add_row(data)
    end

    # Excel-compatible csv, which is a little different to standard csv.
    def to_excel_csv(separator = ',')
      sheet_data.to_csv(separator)
    end

    private def old_range_to_csv(range: dimension, separator: ',')
      csv = CSV.new '', col_sep: separator, quote_char: ?'
      range.each_rowi do |rowix|
        cell_map = row_cell_nodes_at Location[range.top_left.coli, rowix]
        csv << cell_map[range.top_left.coli..range.bot_rite.coli].each.with_index.map{|cell_node, colix| cell_of(cell_node, colix, rowix).formatted_value}
      end
      csv.string
    end

    def range_to_csv(range: dimension, separator: ',')
      csv = CSV.new '', col_sep: separator, quote_char: ?'
      cell_nodes_of(range: range, &method(:cell_of)).map do |row_ary|
        csv << row_ary.map(&:formatted_value)
      end
      csv.string
    end

    # alias to_csv to_excel_csv
    def to_csv(separator = ?,)
      range_to_csv range: dimension, separator: separator
    end

    # create and add a sheet node
    def self.add_node(parent_node, name, sheet_id, relationship_id)
      sheet_node = parent_node.document.create_element("sheet")
      parent_node.add_child(sheet_node)
      sheet_node["name"] = name
      sheet_node["sheetId"] = sheet_id.to_s
      sheet_node["r:id"] = relationship_id
      sheet_node
    end

    private def dimension_node
      # this is about 1.2 - 3 times faster, but optimising this is not worthwhile here.
      # node.children.first.children.find{|n| n.name == 'dimension'}
      node.xpath('xmlns:worksheet/xmlns:dimension').first
    end

    # fetch dimension from the xlsx doc
    def dimension
      @dimension ||= @dimension_fn[]
    end

    # copy dimension data from @dimension to xlsx node
    def update_dimension_node
      if dimension.to_s != dimension_node[:ref]
        dimension_node[:ref] =
        if dimension.count == 1
          'A1'
        else
          dimension.to_s
        end
      end
    end

    private def dimension_of_xlsx
      dimension_str =
      if dimension_node[:ref].include?(?:)
        dimension_node[:ref]
      else
        # sometimes for blank worksheets, dimension_node[:ref] == 'A1'
        # so just make a unit-sized range
        "#{dimension_node[:ref]}:#{dimension_node[:ref]}"
      end

      Office::Range.new dimension_str
    end

    # calculate the actual dimension from the row and cell nodes
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

      if min == Office::Location.largest && max == Office::Location.smallest
        # neither of the reduce ops found anything, so we have a blank worksheet
        Office::Range.new 'A1:A1'
      else
        Office::Range.new min, max
      end
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
        raise LocatorError, "wut!? do what with what rows where!? #{obj.inspect}"
      end
    end

    # Insert new empty rows just before the specified location/range.
    #
    # arg is_a Office::Range (several rows, or one row), or is_a Office::Location (insert 1 row)
    #
    # New row nodes will be returned. As an Array, not as a NodeSet.
    #
    # New row nodes will have only the r= attribute.
    #
    # invalidates row cache, because obviously
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
      # TODO this will break formulas and ranges and drawings referring to these cells
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

    # return the set of ranges specified as merged cells in the xlxs doc
    def merge_ranges
      merge_cells.children.map{|n| Office::Range.new n[:ref]}
    end

    def delete_merge_range range
      # find node and delete it
      merge_cells_node, = node.nspath 'xmlns:worksheet/xmlns:mergeCells'
      to_delete = merge_cells_node.children.find{|n| n[:ref] == range.to_s}
      to_delete or raise LocatorError, "no range found matching #{range}"
      to_delete.unlink
      merge_cells_node[:count] = merge_cells_node.children.count
      to_delete
    end

    # TODO perhaps allow caching of ranges and/or sets, to get the access
    # pattern from higher-level usages.

    # auto-caches row, therefore much faster for things that sequentially access
    # several cells in a row.
    # returns a hash of coli => cell_node for each row
    private def row_cell_nodes_at loc
      @row_cells ||= Array.new
      @row_cells[loc.rowi] ||= begin
        # avoid binding loc into the hash block which would prevent garbage collection
        rowix = loc.rowi

        # When a caller indexes colix, populate the full row.
        # No point waiting for future requests because row_node.element_children
        # iterates from the beginning each time.
        Hash.new do |ha, colix|
          if row_node = row_node_ix(rowix)
            # populate cells for the entire row, and return the cell at colix (if it exists)
            row_node.element_children.reduce nil do |memo_cell,cell_node|
              loc = Location.new(cell_node[:r])
              ha[loc.coli] = cell_node

              # Return value from the Hash.new block must be the cell at colix.
              # So make that the return value from the reduce block.
              if loc.coli == colix
                cell_node
              else
                memo_cell
              end
            end
          end
        end
      end
    end

    # lookup of last row of only 74 rows is around 1ms, cached fetch is around 5-10us
    # so around 100 - 200 times faster.
    def row_node_at loc
      row_node_ix loc.rowi, loc.row_r
    end

    def preload_rows
      @row_nodes ||= begin
        ary = Array.new
        data_node.nxpath("*:row").each do |row_node|
          @row_nodes[row_node[:r].to_i-1] = row_node
        end
        ary
      end
    end

    # fetch node from xml with r=row_r and cache it at rowix
    private def row_node_ix rowix, row_r = rowix+1
      @row_nodes ||= Array.new
      @row_nodes[rowix] ||= begin
        row_node, = data_node.xpath("xmlns:row[@r=#{row_r}]")
        row_node
      end
    end

    # Pre-fetch a range of cell nodes, as an array of arrays.
    # Call optional blk with (colix, rowix, cell_node), useful for constructing a LazyCell or similar
    # will yield nil if no cell found at a location
    def cell_nodes_of range: dimension, &blk
      blk ||= -> i,_c,_r {i} # identity if not specified. Slowdown compared to plain value is microseconds at x10000 repetitions

      range.each_rowi.map do |rowix|
        cell_nodes_map = row_cell_nodes_at(Location[0,rowix])
        range.each_coli.map do |colix|
          blk.call cell_nodes_map[colix], colix, rowix
        end
      end
    end

    # yield a set of enumerators (rows), each of which yields a cell node along
    # with col,row indexes as [colix,rowix]
    # will yield nil if no cell found at a location
    def lazy_cell_nodes_of range: dimension, &row_blk
      return enum_for :lazy_cell_nodes_of, range: range unless block_given?

      # TODO can be more efficient than this, because each cell requires a hash
      # lookup. Which is fast in ruby, but not as fast as a straightforward
      # iteration of row_node.element_children
      range.each_rowi do |rowix|
        # have to construct this with an Enumerator, otherwise 'yield' calls
        # row_blk with the cell. Which is obvs not correct.
        cell_enum = Enumerator.new do |yielder|
          # Use a hash because row children are not always in the correct cell order ...
          cell_nodes_map = row_cell_nodes_at(Location[0,rowix])

          # and look them up from the range
          range.each_coli.each do |colix|
            yielder.yield cell_nodes_map[colix], colix, rowix
          end
        end
        yield cell_enum
      end
    end

    # return an Array of Arrays of the specified range, with optional tranformation block
    #
    # eg sheet.cells_of(Range.new('B17:F23'), &:formatted_value)
    def cells_of range = dimension, &blk
      blk ||= ->i{i} # identity if not specified. Slowdown compared to plain value is microseconds at x10000 repetitions

      @cells ||= {}
      ary_of_arys = lazy_cell_nodes_of(range: range).map do |row_enum|
        row_enum.map do |cell_node, colix, rowix|
          cell = @cells[[colix, rowix]] ||= cell_of cell_node, colix, rowix
          blk[cell]
        end
      end
      ary_of_arys
    end

    private def strict_cell_of cell_node, colix = nil, rowix = nil
      case [cell_node, colix, rowix].map(&:class)
      when [Nokogiri::XML::Node, NilClass, NilClass]
        Cell.new cell_node, workbook.shared_strings, workbook.styles

      when [NilClass, Integer, Integer]
        LazyCell.new self, Location[colix, rowix]

      # TODO in this case we assume that caller has verified that colix and rowix are correct.
      when [Nokogiri::XML::Node, Integer, Integer]
        Cell.new cell_node, workbook.shared_strings, workbook.styles
      end
    end

    # Mind the footguns.
    private def loose_cell_of cell_node, colix = nil, rowix = nil
      if cell_node
        Cell.new cell_node, workbook.shared_strings, workbook.styles
      else
        LazyCell.new self, Location[colix, rowix]
      end
    end

    # Create a cell from either cell_node, or [colix,rowix]
    private def cell_of cell_node, colix = nil, rowix = nil
      @cells ||= {}
      case
      when colix && rowix
        @cells[[colix, rowix]] ||= loose_cell_of cell_node, colix, rowix

      when cell_node
        cell = loose_cell_of cell_node
        @cells[[cell.location.coli, cell.location.rowi]] ||= cell

      else
        raise Error, "cannot build cell from #{{cell_node: cell_node, colix: colix, rowix: colix}}"
      end
    end

    # TODO we could be slightly smarter than throwing the whole cache away, but
    # that's a quite lot more code.
    def invalidate_row_cache
      @cells = {}

      @dimension = nil
      # take dimension from the rows and cells now
      @dimension_fn = method :calculate_dimension

      @row_cells = Array.new
      @row_nodes = Array.new
      @sheet_data = nil
    end

    # get a cell at the given location, which may be Location, [Integer,Integer], or "A1"
    def [](*args)
      # will work better as case .. in with ≥ruby-2.7
      case args.map(&:class)
      when [Integer, Integer]
        coli, rowi = args
        self[ Location[coli, rowi] ]

      when [String]
        a1_location, = args
        self[ Location.new(a1_location) ]

      when [Location]
        loc, = args
        cell_node = row_cell_nodes_at(loc)&.dig(loc.coli)
        cell_of cell_node, *loc

      else
        raise LocatorError, "don't know how to get cell from #{args.inspect}"
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
      data_node.nspath('~row/~c').each do |c_node|
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

    # this is actually a document node
    def node; worksheet_part.xml end
    def to_xml; node.to_xml end

    def inspect
      lazy_cell_nodes_of.map do |row|
        row.map do |cell_node, colix, rowix|
          cell = cell_of cell_node, colix, rowix
          cell.formatted_value || cell.location.to_s.to_sym
        end
      end
    end

    # Not currently used, but could obviate creation of a LazyCell instance
    # Difference that .cell = could be lazy, whereas this could be immediate.
    # Not sure if that makes any sense.
    def []=(coli,rowi,value)
      case sheet_data.rows
      when NilClass
        LazyCell.new(self, coli, rowi).value = value
      else
        rowi.cells[coli].value = value
      end
    end

    # Fill range with corresponding values from data. Assumes that
    # range.top_left should be set to data[0][0]
    #
    # data must provide #[] for itself and its elements. Array of Arrays would work.
    #
    # returns range
    #
    # Existing values will be overwritten. New cells will be created on-demand.
    def project!(range, data)
      range.each_by_row do |colix, rowix|
        self[colix,rowix].value = data[rowix - range.top_left.rowi][colix - range.top_left.coli]
      end
      invalidate_row_cache
      range
    end

    # Accept data into its rectangle with location as top-left.
    #
    # data must provide each_with_index, so it should probably be an Enumerable
    # of Enumerables. Array of Arrays would work.
    #
    # returns the range filled
    #
    # Existing values will be overwritten. New cells will be created on-demand.
    def accept!(location, data)
      furthest = location.dup
      data.each_with_index do |row,rowix|
        row.each_with_index do |value,colix|
          cell_location = location + [colix, rowix]

          # Apparently there really needs to be a unified way to set cell
          # 'values' even when the value is an image. Although we don't have
          # extent here, so what size to use?
          self[cell_location].value = case value
          when Magick::ImageList, Magick::Image
            value.inspect
          else
            value
          end

          furthest |= cell_location
        end
      end
      invalidate_row_cache
      Range.new location, furthest
    end

    # fetch the drawing part for this sheet
    def drawing_part
      @drawing_part ||= fetch_drawing_part || create_drawing_part
    end

    # 15-Apr-2021 so far this only exists so that external code can find out
    # whether image creation works. Which is better than having external code
    # digging far into Sheet internals.
    def has_drawing?
      !!drawing_rel_id
    end

    private def drawing_rel_id
      # Nokogiri::XML::Element does not understand #dig
      if drawing_node = node.nxpath('*:worksheet/*:drawing').first
        drawing_node['r:id']
      end
    end

    private def fetch_drawing_part
      if dri = drawing_rel_id
        rel_part = worksheet_part.get_relationship_by_id dri
        rel_part.target_part
      end
    end

    # From OfficeOpenXML-XMLSchema-Strict.zip/sml.xsd/xsd:complexType[@name="CT_Worksheet"]
    # Hash of tag name to integer order
    SHEET_CHILD_NODE_ORDER = (<<~EOTS).split(/\s+/).each_with_index.each_with_object({}){|(name,index),ha| ha[name] = index }
      sheetPr
      dimension
      sheetViews
      sheetFormatPr
      cols
      sheetData
      sheetCalcPr
      sheetProtection
      protectedRanges
      scenarios
      autoFilter
      sortState
      dataConsolidate
      customSheetViews
      mergeCells
      phoneticPr
      conditionalFormatting
      dataValidations
      hyperlinks
      printOptions
      pageMargins
      pageSetup
      headerFooter
      rowBreaks
      colBreaks
      customProperties
      cellWatches
      ignoredErrors
      smartTags
      drawing
      drawingHF
      picture
      oleObjects
      controls
      webPublishItems
      tableParts
      extLst
    EOTS

    # Sort child nodes in the order specified by sml.xsd, otherwise Excel throws
    # its toys.
    private def fixup_drawing_tag_order
      # Sort non-text nodes in the right order. Unknown node names at the end.
      sorted_child_node_ary = node.root.element_children.sort_by{|n| SHEET_CHILD_NODE_ORDER[n.name] || Float::INFINITY}

      # Unlink all child nodes ...
      child_ary = node.root.children.unlink

      # ... then reattach iteratively in the correct order while respecting
      # text nodes (ie whitespace).
      child_ary.reduce 0 do |node_index, child_node|
        if child_node.text?
          node.root << child_node
          node_index
        else
          node.root << sorted_child_node_ary[node_index]
          node_index + 1
        end
      end
    end

    # create, attach and return a drawing part
    private def create_drawing_part
      # create drawing
      drawing = ImageDrawing.build_wsdr.doc

      # drawing part added to workbook as drawings/drawingX ...
      drawing_part = workbook.add_drawing_part(drawing, workbook.workbook_part.path_components)

      # ... with rel from sheet to drawing
      drawing_rel_id = workbook.add_relationship(worksheet_part, drawing_part, DRAWING_RELATIONSHIP_TYPE)

      # append the <drawing r:id="rIdX"/> node as a child of worksheet
      Nokogiri::XML::Builder.with node.nxpath('*:worksheet').first do |bld|
        bld.drawing 'r:id': drawing_rel_id
      end

      fixup_drawing_tag_order

      drawing_part
    end

    # fetch the wsDr node in the drawing (which all images are attached to)
    def drawing_wsdr_node
      drawing_part.xml.nxpath('*:wsDr').first
    end

    # add the image to display anchored at loc, with optional width x height
    # return the drawing part containing the image
    # TODO use stretch if placeholder extent is larger than native image
    # TODO inserting rows/columns would break the drawing data.
    # MAYBE where does image_name come from File.basename(image_part.name) would work I think
    def add_image(image, loc, extent: nil)
      # image part added to workbook as media/imageX
      image_part = workbook.add_image_part(image, workbook.workbook_part.path_components)

      # create anchor in drawing (with unique-enough tmp value in blip@r:embed)
      tmp_rel_id = "tmp_rel_id-#{Base64.urlsafe_encode64 "#{Time.now}/#{Thread::current.__id__}"}"
      image_drawing = ImageDrawing.new img: image, loc: loc, rel_id: tmp_rel_id, extent: extent
      image_drawing.build_anchor drawing_wsdr_node

      # ... and rel from drawing -> image
      # TODO does this still apply?
      image_rel_id = workbook.add_relationship(drawing_part, image_part, IMAGE_RELATIONSHIP_TYPE)

      # Using tmp value update r:embed attribute to rId from the image rel.
      # Remember that drawing might have been copied so we need to update the
      # active one, ie drawing_part.xml
      # Also, the r:embed attribute needs the r namespace to be declared. If
      # this code constructed the drawing's wsDr tag it will be fine, but if something
      # else did then it may have optimised the namespaces and removed r. So don't
      # assume that r:embed will work properly in an xpath.
      # Also also, I suspect that a namespace being declared in the same node
      # where an attribute refers to that namespace might tickle a bug in
      # nokogiri or libxml2.
      blip_node, = drawing_part.xml.nxpath(%|/*:wsDr/*:oneCellAnchor/*:pic/*:blipFill/*:blip[@*:embed = '#{tmp_rel_id}']|)

      # Assuming there will only be one namespace for embed.
      blip_node.attributes['embed'].value = image_rel_id

      image_part
    end
  end
end
