module Word
  module Options
    class HideMap < Word::Option
      def importance
        6
      end

      def apply_option
        if params.downcase == 'true' && placeholder.field_value.length > 1
          placeholder.field_value = [placeholder.field_value[1]]
        end
      end
    end
  end
end
