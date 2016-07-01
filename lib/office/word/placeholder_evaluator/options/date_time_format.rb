module Word
  module Options
    class DateTimeFormat < Word::Option
      def apply_option
        value = placeholder.field_value
        if !value.nil? && value.length > 0
          begin
            date_time_string = DateTime.parse(value.to_s).strftime(params)
          rescue
            date_time_string = parse_arabic(value.to_s).strftime(params)
            date_time_string = numbers_to_arabic(date_time_string)
          end
          placeholder.field_value = date_time_string
        end
      end

      def parse_arabic(value)
        date_time = DateTime.parse(arabic_to_numbers(value))
      end

      def arabic_to_numbers(text)
        text.tr('٠١٢٣٤٥٦٧٨٩', '0123456789')
      end

      def numbers_to_arabic(text)
        text.tr('0123456789', '٠١٢٣٤٥٦٧٨٩')
      end
    end
  end
end
