require_relative 'placeholder_grammar.rb'
require_relative 'placeholder_lexer.rb'

module Office
  # for parsing placeholder text and handing back results
  class Placeholder
    def initialize field_path, options
      @field_path, @raw_options = field_path, options
    end

    attr_reader :field_path

    def self.rejoin parts
      parts.reduce '' do |str, part|
        case part
        when Integer
          str << ?[ << part.to_s << ?]
        when Symbol
          str << ?. unless str.empty?
          str << part.to_s
        else
          raise "unknown part in path: #{part.inspect}"
        end
      end
    end

    def expr
      @expr = self.class.rejoin field_path
    end

    def options
      @options ||= @raw_options[:keywords].merge @raw_options[:functors]
    end

    def image_extent
      @raw_options[:image_extent]
    end

    # placeholder_str is things like
    # entries.your_picture|100x200
    # entries.your_picture|size(100,200)
    # entries.entries.group|layout(d3:g4)
    def self.parse placeholder_str
      grammar = Office::PlaceholderGrammar.new
      grammar.read_tokens Office::PlaceholderLexer.tokenize placeholder_str
      values = grammar.to_h
      field_path = values.delete :field_path
      new field_path, values
    end
  end
end
