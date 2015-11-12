module Word
  module Options
    class ShowCoordinateInfo < Word::Option
      def importance
        5
      end

      def apply_option
        if params.downcase == 'false'
          placeholder.field_value = [placeholder.field_value[0]]
        end
      end
    end
  end
end
