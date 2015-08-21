module Word
  module GroupOptions
    class List < Word::GroupOption
      def apply_option
        placeholder.group_generation_options[:generation_method] = :list
      end
    end
  end
end
