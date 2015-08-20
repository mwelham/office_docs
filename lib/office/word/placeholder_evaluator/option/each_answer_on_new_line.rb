class PlaceholderEvaluator::Option::EachAnswerOnNewLine
  def apply_option
    placeholder.field_value = placeholder.field_value.split(',').map(&:strip)
  end
end
