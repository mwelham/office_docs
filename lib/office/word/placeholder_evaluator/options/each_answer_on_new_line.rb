module Word
  module Options
    class EachAnswerOnNewLine < Word::Option
      def apply_option
        placeholder.field_value = placeholder.field_value.split(',').map(&:strip).join("\n")
      end
    end
  end
end
