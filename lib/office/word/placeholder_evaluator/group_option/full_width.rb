class PlaceholderEvaluator::GroupOption::FullWidth
  def apply_option
    placeholder.group_generation_options[:generation_method] = :table
    placeholder.render_options[:use_full_width] = true
  end
end
