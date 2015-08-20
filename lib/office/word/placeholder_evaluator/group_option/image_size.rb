class PlaceholderEvaluator::GroupOption::ImageSize
  def apply_option
    width, height = get_width_and_height_from_params
    placeholder.group_generation_options[:image_size][:width] = width
    placeholder.group_generation_options[:image_size][:height] = height
  end
end
