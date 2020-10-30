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
