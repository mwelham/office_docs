require 'date'

require_relative 'cell.rb'

module Office
  # Intended as a placeholder for a cell, but does not add nodes to the xml
  # until it's given a value.
  class LazyCell
    include CellNodes

    def initialize sheet, *args
      raise "Not a sheet" unless sheet.is_a? Sheet
      @sheet = sheet

      case args.map(&:class)
      when [Integer, Integer]
        coli,rowi = args
        Location[coli,rowi]
      when [Location]
        loc, = args
        @location = loc
      else
        raise "dunno how to construct #{self.class} from #{args.inspect}"
      end
    end

    def styles
      sheet.workbook.styles
    end

    attr_reader :sheet, :location
    private :sheet

    def empty?; true end

    # always nil
    def value; end
    def formatted_value; end

    def value=(obj)
      # TODO could maybe possibly optimise this using the row/@r numbers and row[position() = offset]
      #   using the sheet dimension to calculate offset
      #
      # TODO could optimise by storing the row node in the lazy cell on
      # creation, since anyway that part of the node has to check whether the
      # row exists.
      row_node = sheet.row_node_at location
      if row_node.nil?
        # create row_node, then add cell in appropriate place
        row_node, = sheet.insert_rows location
      end

      # create c node and set its value
      c_node = CellNodes.build_c_node \
        sheet.node.document.create_element(?c, r: location.to_s),
        obj,
        styles: styles

      # TODO can we always just add to the end of the c children, or must they be in r order?
      row_node << c_node

      # TODO forward future calls to real cell?
      # or maybe have a module. It's the usual "change the class of an instance from the inside" problem
    end
  end
end
