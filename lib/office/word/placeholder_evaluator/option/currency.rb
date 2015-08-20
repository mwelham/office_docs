module Word
  module Options
    class Currency < Word::Option
      def apply_option
        placeholder.field_value = number_to_currency(placeholder.field_value.to_f, unit: '')
      end
    end
  end
end
