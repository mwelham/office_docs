module Office
  # Algebra for A1 style cell locations. Well, Geometry really cos it's 2d.
  #
  # Convert to/from strings and [0,0] numeric indices
  class Location
    # constructor for 'A1' style strings
    # col_row_fns is for internal use only.
    #
    # NOTE you could pass in a garbage location_string and nothing would go
    # wrong until you tried to do something other than to_s.
    def initialize location_string, col_row_fns: nil
      @location_string = location_string&.dup&.freeze
      @col_row_fns = col_row_fns
    end

    # Constructor for 2x integers. Should really be positive integers, otherwise
    # weird things will happen.
    def self.[](coli, rowi)
      # make sure these meet the type constraints, fail early if not
      coli, rowi = Integer(coli), Integer(rowi)
      new nil, col_row_fns: [->{coli}, ->{rowi}]
    end

    # This constructor is special-purpose for constructing
    # from the results of a call to .parse_a1, because it's useful
    # when working directly with <row r=> and <col r=>
    #
    # colst must be [A-Z]+
    # rowst must be 1-based and can be Integer or String
    def self.of_r colst, rowst
      # make sure these meet the type constraints, fail early if not
      rowi = Integer rowst
      coli = Integer col_index(colst.to_s)

      new nil, col_row_fns: [->{coli}, ->{rowi - 1}]
    end

    # useful for reduce and other operations requiring a unit element
    def self.unit
      new nil, col_row_fns: [->{0}, ->{0}]
    end

    def location_string
      @location_string ||= "#{self.class.column_name(coli)}#{rowi+1}"
    end

    # Union of two ranges.
    #
    # Result is the supremum (ie max of both dimensions) of the arguments
    #
    # A5 | A6 => A6
    # A5 | B4 => B5
    # A8 | C1 => C8
    def | rhs
      # TODO could optimise by only creating a new instance when it's different to both self and rhs
      # OR have a giant hash of [coli, rowi] => Location, since these are really value objects
      self.class[ [self.coli,rhs.coli].max, [self.rowi,rhs.rowi].max ]
    end

    # Intersection of two ranges.
    #
    # Result is the infimum (ie min of both dimensions) of the arguments
    #
    # A5 & A6 => A5
    # A5 & B4 => A4
    # A8 & C1 => A1
    def & rhs
      # TODO could optimise by only creating a new instance when it's different to both self and rhs
      # OR have a giant hash of [coli, rowi] => Location, since these are really value objects
      self.class[ [self.coli,rhs.coli].min, [self.rowi,rhs.rowi].min ]
    end

    def coli; @coli ||= col_row_fns.first.call end
    def rowi; @rowi ||= col_row_fns.last.call end

    alias to_s location_string
    def to_a; [coli, rowi] end
    def inspect; "#{to_s}(#{coli},#{rowi})" end

    # this is the 1-based index for using with <row r="#{row_r}"/>
    def row_r; rowi + 1 end

    # Comparison. Will also work when rhs is an A1 style string.
    def == rhs
      # to_a would also work. But straight string comparison is probably faster than constructing two temp arrays.
      self.to_s == rhs.to_s
    end

    # match with A1-style location strings
    def === loc
      self.to_s == loc
    end

    # Returns a new Location modified by args,
    # where args is a [column_delta,row_delta] 2-element array
    #
    # example:
    #  loc = Location.new 'B15'
    #  => B15(1,14)
    #  loc + [1,10]
    #  => C25(2,24)
    def + deltas
      case deltas
      in [Integer => cold, Integer => rowd]
        self.class[coli + cold, rowi + rowd]
      else
        raise "#{self.class} dunno how to extend by #{deltas.inspect}"
      end
    end

    # Returns a new Location modified by args,
    # where args is a [column_delta,row_delta] 2-element array
    #
    # example:
    #  loc = Location.new 'D27'
    #  => D27(3,26)
    #  loc - [2,13]
    #  => B14(1,13)
    def - deltas
      # just add the inverse
      self + deltas.map{|i| -i}
    end

    # Lazily create the lazy coli and rowi values. Yes, both are lazy.
    private def col_row_fns
      @col_row_fns ||= begin
        colst, rowst = self.class.parse_a1 location_string
        # Store as lambdas so we don't waste time calculating col needlessly if
        # it wasn't needed.
        #
        # NOTE we assume from parse that colst and rowst are the correct types.
        colfn = ->{self.class.col_index colst}
        rowfn = ->{Integer(rowst) - 1}
        [colfn, rowfn]
      end
    end

    # The intention here is to provide an intermediate parsing layer where you
    # just get the data fields back. Less convenient than using Location.new, but
    # occasionally necessary.
    #
    # eg parse 'CA256' => ["CA", "256"]
    def self.parse_a1 a1_location
      # TODO this will fail with $A$1 and similar absolute references
      /^(?<colst>[[:upper:]]+)(?<rowst>[[:digit:]]+)$/ =~ a1_location or raise "#{self} can't parse '#{a1_location}'"
      return colst, rowst
    end

    A_ORD = ?A.ord

    # Convert colst from digits in base 26 (ie A-Z) to a zero-based integer index. Inverse of column_name.
    #
    # example:
    #  Location.col_index "A"  => 0
    #  Location.col_index "B"  => 1
    #  Location.col_index "IW" => 256
    def self.col_index colst
      base = 26
      # This is just the standard base_conversion algorithm where A => 0, B => 1 etc
      coli_base_one = colst.each_char.reverse_each.each_with_index.reduce 0 do |ax,(digit,position)|
        ax + (digit.ord - A_ORD + 1) * base ** position
      end

      coli_base_one - 1
    end

    # convert zero-based integer index to column A-Z. Inverse of col_index.
    #
    # example:
    #  Location.column_name 0   => "A"
    #  Location.column_name 1   => "B"
    #  Location.column_name 256 => "IW"
    def self.column_name(index)
      # copypasta'd from old version of Cell
      name = ''
      while index >= 0
        name << (A_ORD + (index % 26)).chr
        index = index/26 - 1
      end
      name.reverse
    end
  end
end
