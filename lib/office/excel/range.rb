module Office
  # Algebra for B4:H22 rectangular ranges
  class Range
    attr_reader :top_left
    attr_reader :bot_rite

    def initialize *args
      case args
      in [String => range_str]
        # construct from string
        @range_str = range_str.dup.freeze
        @top_left, @bot_rite = range_str.split(':').map{|loc_str| Location.new loc_str}
      in [Location => tl, Location => bt]
        @top_left, @bot_rite = tl.dup, bt.dup
      else
        raise "dunno how to create #{self.class} from #{args.inspect}"
      end
    end

    # where location is an Office::Location
    def cover? location
      (top_left.coli..bot_rite.coli).cover?(location.coli) \
      and \
      (top_left.rowi..bot_rite.rowi).cover?(location.rowi)
    end

    def width
      @bot_rite.coli - top_left.coli + 1
    end

    def height
      @bot_rite.rowi - top_left.rowi + 1
    end

    # yield each row_r to blk
    #
    # NOTE 1-based not zero-based
    def each_row_r &blk
      return enum_for :each_row_r unless block_given?
      (top_left.row_r..bot_rite.row_r).each &blk
    end

    # yield each col index to blk
    def each_coli &blk
      return enum_for :each_coli unless block_given?
      (top_left.coli..bot_rite.coli).each &blk
    end

    def each_by_row &blk
      return enum_for :each_by_row unless block_given?

      (top_left.rowi..bot_rite.rowi).each do |rowix|
        (top_left.coli..bot_rite.coli).each do |colix|
          yield Location[colix,rowix]
        end
      end
    end

    def each_by_col &blk
      return enum_for :each_by_col unless block_given?

      (top_left.coli..bot_rite.coli).each do |colix|
        (top_left.rowi..bot_rite.rowi).each do |rowix|
          yield Location[colix,rowix]
        end
      end
    end

    def to_s
      @range_str ||= "#{top_left.to_s}:#{bot_rite.to_s}"
    end

    # will also compare to string representations
    def == rhs
      self.to_s == rhs.to_s
    end

    def to_a; [top_left, bot_rite]; end

    def inspect; [top_left.inspect, bot_rite.inspect].inspect; end
  end
end
