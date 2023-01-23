module Word
  module Options
    class Separator < Word::Option
      def apply_option
        separator = params
        placeholder.field_value = placeholder.field_value.split(',').map(&:strip).join(separator)
      end
    end
  end
end
