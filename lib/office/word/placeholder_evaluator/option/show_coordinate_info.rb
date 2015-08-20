class PlaceholderEvaluator::Option::ShowCoordinateInfo
  def apply_option
    if params.downcase == 'false'
      placeholder.field_value = [placeholder.field_value[0]]
    end
  end
end
