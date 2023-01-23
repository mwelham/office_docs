module Word
  module Options
    class ImageSize < Word::Option
      def apply_option
        if placeholder.is_image_answer?
          width, height = get_width_and_height_from_params
          if resample?
            placeholder.field_value = resize_image_answer(placeholder.field_value, width, height)
          else
            placeholder.render_options = image_constraints(placeholder.field_value, width, height)
          end
        end
      end
    end
  end
end
