module Word
  module Options
    class EachAnswerOnNewLine < Word::Option
      def apply_option
        placeholder.field_value = placeholder.field_value.split(',').map(&:strip)
      end
    end
  end
end
