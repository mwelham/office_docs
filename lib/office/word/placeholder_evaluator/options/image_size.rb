module Word
  module Options
    class ImageSize < Word::Option
      def apply_option
        width, height = get_width_and_height_from_params
        placeholder.field_value = resize_image_answer(placeholder.field_value, width, height)
      end
    end
  end
end
