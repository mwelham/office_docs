module Word
  module Options
    class DateTimeFormat < Word::Option
      def apply_option
        date_time = DateTime.parse(placeholder.field_value.to_s)
        placeholder.field_value = date_time.strftime(params)
      end
    end
  end
end
