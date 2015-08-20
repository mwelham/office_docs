class PlaceholderEvaluator::Option::MapSize
  def apply_option
    width, height = get_width_and_height_from_params
    placeholder.field_value[0] = resize_image_answer(placeholder.field_value[0], width, height)
  end
end
