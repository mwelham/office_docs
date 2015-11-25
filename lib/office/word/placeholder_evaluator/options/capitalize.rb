module Word
  module Options
    class Capitalize < Word::Option
      def apply_option
        placeholder.field_value = placeholder.field_value.capitalize if placeholder.field_value.is_a?(String)
      end
    end
  end
end
