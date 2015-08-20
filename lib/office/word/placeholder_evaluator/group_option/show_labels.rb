class PlaceholderEvaluator::GroupOption::ShowLables
  def apply_option
    if params.downcase == 'false'
      placeholder.group_generation_options[:show_labels] = false
      placeholder.render_options[:no_header_row] = true
    end
  end
end
