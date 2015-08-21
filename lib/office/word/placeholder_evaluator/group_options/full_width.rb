module Word
  module GroupOptions
    class FullWidth < Word::GroupOption
      def apply_option
        placeholder.group_generation_options[:generation_method] = :table
        placeholder.render_options[:use_full_width] = true
      end
    end
  end
end
