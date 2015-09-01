module Word
  module Options
    class DateTimeFormat < Word::Option
      def apply_option
        value = placeholder.field_value
        unless value.nil?
          date_time = DateTime.parse(value.to_s)
          placeholder.field_value = date_time.strftime(params)
        end
      end
    end
  end
end
