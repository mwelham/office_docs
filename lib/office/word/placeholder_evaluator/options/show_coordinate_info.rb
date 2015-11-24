module Word
  module Options
    class ShowCoordinateInfo < Word::Option
      def importance
        5
      end

      def apply_option

        if params.downcase == 'false' && placeholder.field_value.length > 1
          placeholder.field_value = [placeholder.field_value[0]]
        elsif params != 'true'
          # some kind of formatting
          # TODO - some kind of template replace
          # e.g. %lat, %long

          # index = placeholder.field_value.length > 1 ? 1 : 0
          # new_coord_info = get_new_coord_info_from_template(placeholder.field_value[index], params)
          # placeholder.field_value[index] = new_coord_info
        end
      end
    end
  end
end
