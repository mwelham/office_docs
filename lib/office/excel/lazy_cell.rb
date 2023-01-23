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

    attr_reader :sheet, :location
    private :sheet

    # never a placeholder
    def placeholder; end

    def empty?; true end

    # always nil
    def value; end
    def to_ruby; end
    def formatted_value; end

    def value=(obj)
      sheet[location] = obj
      # TODO forward future calls to real cell?
      # or maybe have a module. It's the usual "change the class of an instance from the inside" problem
    end
  end
end
