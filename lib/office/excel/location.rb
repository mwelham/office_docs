module Office
  # Algebra for A1 style cell locations. Well, Geometry really cos it's 2d.
  #
  # Convert to/from strings and [0,0] numeric indices
  class Location
    # constructor for 'A1' style strings
    # col_row_fns is for internal use only.
    def initialize location_string, col_row_fns: nil
      @location_string = location_string&.dup&.freeze
      @col_row_fns = col_row_fns
    end

    # constructor for 2x integers
    def self.[](coli,rowi)
      new nil, col_row_fns: [->{coli}, ->{rowi}]
    end

    # useful for reduce and other operations requiring a unit element
    def self.unit
      new nil, col_row_fns: [->{0}, ->{0}]
    end

    def location_string
      @location_string ||= "#{self.class.column_name(coli)}#{rowi+1}"
    end

    # result is the supremum of the two dimensions of each argument
    # A5 | A6 => A6
    # A5 | B4 => B5
    # A8 | C1 => C8
    def | rhs
      # TODO could optimise by only creating a new instance when it's different to both self and rhs
      # OR have a giant hash of [coli, rowi] => Location, since these are really value objects
      self.class[ [self.coli,rhs.coli].max, [self.rowi,rhs.rowi].max ]
    end

    # result is the infimum of the two dimensions of each argument
    # A5 & A6 => A5
    # A5 & B4 => A4
    # A8 & C1 => A1
    def & rhs
      # TODO could optimise by only creating a new instance when it's different to both self and rhs
      self.class[ [self.coli,rhs.coli].min, [self.rowi,rhs.rowi].min ]
    end

    def to_s; location_string end
    def coli; @coli ||= col_row_fns.first.call end
    def rowi; @rowi ||= col_row_fns.last.call end
    def to_a; [coli, rowi] end
    def inspect; "#{to_s}(#{coli},#{rowi})" end

    # this is the 1-based index for using with <row r="#{row_r}"/>
    def row_r; rowi + 1 end

    def == rhs
      # to_a would also work. But straight string comparison is probably faster than constructing two temp arrays.
      self.to_s == rhs.to_s
    end

    # match with A1-style location strings
    def === loc
      self.to_s == loc
    end

    # args is a [column_delta,row_delta] 2-element array
    def + args
      case args
      in [Integer => cold, Integer => rowd]
        self.class[coli + cold, rowi + rowd]
      else
        raise "#{self.class} dunno how to extend by #{args.inspect}"
      end
    end

    A_ORD = ?A.ord - 1

    private def col_row_fns
      @col_row_fns ||= begin
        # TODO this will fail with $A$1 and similar absolute references
        /^(?<colst>[[:alpha:]]+)(?<rowst>[[:digit:]]+)$/ =~ location_string.upcase
        # treat chars as digits in base 26, and store as lambdas so we don't
        # waste time calculating col needlessly if it wasn't needed. Kinda pointless optimisation. But anyway. Not like it was hard to put it in a lambda :-}
        colfn = ->{colst.each_char.reverse_each.each_with_index.reduce(0){|s,(ch,i)| s + (ch.ord - A_ORD) ** (i+1)} - 1}
        rowfn = ->{Integer(rowst) - 1}
        [colfn, rowfn]
      end
    end

    def self.column_name(index)
      # copypasta'd from old version of Cell
      name = ''
      while index >= 0
        name << ('A'.ord + (index % 26)).chr
        index = index/26 - 1
      end
      name.reverse
    end
  end
end
