require_relative 'lexer_error_info'

module Office
  # It should be possible to adapt this lexer to read out of a nokogiri doc and parse the placeholders like that.
  # doc.nxpath('//*:p/text()')
  module PlaceholderLexer
    using LexerErrorInfo

    # Only difference is the kind of quote matched
    # mostly borrowed from some json lexer
    DQUOTE_RX = /"([^"\\]|\\["\\\/bfnrt])*?"/
    SQUOTE_RX = /'([^'\\]|\\['\\\/bfnrt])*?'/
    LRQUOTE_RX = /[“”]([^'\\]|\\[“”\\\/bfnrt])*?[“”]/

    # The lexer.
    # yield (symbol, value) pairs where value is extended with
    # some lexer info.
    def self.tokenize line
      return enum_for __method__, line unless block_given?
      s = StringScanner.new line
      # Do a workaround here and attach the current pos and string ref
      # to every value yielded, so if an error occurs those can be used
      # for making nice error messages.
      nopos_tokenize s do |token, value|
        yield token, value.spos(s)
      end
    end

    # yield plain value, ie no special lexer info
    # s is a Scanner
    def self.nopos_tokenize s
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
        # but causes some trouble for functor(a3:g7) because of the leading :
        # so this is not really a context-free grammar.
        when s.scan(/:(?>\s*)([^'"“”\[].*?)\s*([,})])/)
          # yield : to not break the grammar too much
          yield ?:, ?:

          # try to reuse existing lexing
          subtokens = Array tokenize s.captures[0]
          if subtokens.size == 1 && subtokens.first.first.is_a?(Symbol)
            yield *subtokens.first
          else
            # otherwise yield this as a special kind of STRING, which includes naked punctuation
            # and space-separated stuff, and weirdly-quoted stuff.
            yield :MAGIC_QUOTED, s.captures[0]
          end
          # this steals , } ) from the rest of the lexer, so just yield it here
          yield s.captures[1], s.captures[1]

        when s.scan(/\d+/i);       yield :NUMBER, s.matched
        when s.scan(/\w[\d\w]*/); yield :IDENTIFIER, s.matched
        when s.skip(/\s/);         # ignore white space

        # hoop-jumping to match various kinds of quotes
        # TODO consolidate these. There must be a better way to do quote-matching. Maybe regex back-references?
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
