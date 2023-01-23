module LexerErrorInfo
  # Using this otherwise we have to wrap every token value in something. This
  # seems less painful.

  refine String do
    def spos scanner
      @_lexer_pos = (scanner.pos || 0) - self.length
      @_lexer_string = scanner.string
      self
    end

    def lexer_pos
      @_lexer_pos
    end

    def lexer_string
      @_lexer_string
    end
  end
end
