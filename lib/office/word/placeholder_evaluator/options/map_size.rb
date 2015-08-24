module Word
  module Options
    class MapSize < Word::Option
      def apply_option
        width, height = get_width_and_height_from_params
        return true if placeholder.field_value.nil?
        placeholder.field_value[0] = resize_image_answer(placeholder.field_value[0], width, height)
      end
    end
  end
end
