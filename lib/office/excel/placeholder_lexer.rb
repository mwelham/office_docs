module Office
  module PlaceholderLexer
    DQUOTE_RX = /"([^"\\]|\\["\\\/bfnrt])*?"/
    SQUOTE_RX = /'([^'\\]|\\['\\\/bfnrt])*?'/
    LRQUOTE_RX = /[“”]([^'\\]|\\[“”\\\/bfnrt])*?[“”]/

    # The lexer.
    def self.tokenize line
      return enum_for __method__, line unless block_given?

      s = StringScanner.new line
      case
        when s.scan(/true/); yield [:true, 'true']
        when s.scan(/false/); yield [:false, 'false']
        when s.scan(/(\d+)x(\d+)/i)
          yield :NUMBER, s.captures[0]
          yield ?x, ?x
          yield :NUMBER, s.captures[1]

        # hoop-jumping to handle keywords with:
        # - bare symbols eg 'separator: ;'
        # - unquoted values eg 'date_time_format: %d &m %y'
        when s.scan(/:(?>\s*)([^'"“”\[].*?)\s*([,}])/)
          yield ?:, ?:
          yield :MAGIC_QUOTED, s.captures[0]
          yield s.captures[1], s.captures[1]

        when s.scan(/\d+/i);       yield :NUMBER, s.matched
        when s.scan(/\w[\d\w_]*/); yield :IDENTIFIER, s.matched
        when s.skip(/\s/);         # ignore white space

        # hoop-jumping to match various kinds of quotes
        # TODO consolidate these
        when s.scan(SQUOTE_RX)
          str = s.matched
          yield str[0], str[0]
          yield :STRING, s.matched[1...-1]
          yield str[-1], str[-1]

        when s.scan(DQUOTE_RX)
          str = s.matched
          yield str[0], str[0]
          yield :STRING, s.matched[1...-1]
          yield str[-1], str[-1]

        when s.scan(LRQUOTE_RX)
          str = s.matched
          yield :LRQUOTE, str[0]
          yield :STRING, s.matched[1...-1]
          yield :LRQUOTE, str[-1]

        else
          nc = s.getch
          yield nc, nc
      end until s.eos?
    end
  end
end
