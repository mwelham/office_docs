class PlaceholderEvaluator::GroupOption::List
  def apply_option
    placeholder.group_generation_options[:generation_method] = :list
  end
end
