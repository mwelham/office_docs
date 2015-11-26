module Word
  module ForLoopExpanders
    class Base
      attr_accessor :main_doc, :data, :options
      def initialize(main_doc, data, options = {})
        self.main_doc = main_doc
        self.data = data
        self.options = options
      end

      def replace_variable_in_placeholders(index, for_loop_placeholder_info, placeholders, paragraph, inbetween_runs=nil)
        placeholders.each do |p|
          placeholder_variable_matcher = /#{for_loop_placeholder_info[:variable]}\./
          placeholder = p[:placeholder_text]
          if placeholder.match(placeholder_variable_matcher)
            new_placeholder = placeholder.gsub(placeholder_variable_matcher,"#{for_loop_placeholder_info[:data_pointer]}[#{index}].")
            if inbetween_runs
              paragraph.replace_all_with_text(placeholder, new_placeholder, inbetween_runs)
            else
              paragraph.replace_all_with_text(placeholder, new_placeholder)
            end
          end
        end
      end

      def parse_for_loop_placeholder(placeholder)
        result = placeholder.gsub('{%','').gsub('%}','').match(/foreach (\w+) in (.+)/)
        variable = result[1].strip
        data_pointer = result[2].strip
        raise "Invalid syntax for foreach placeholder #{placeholder}" if variable.blank? || data_pointer.blank?
        field_data = Word::Template.get_value_from_field_identifier(data_pointer, data)
        {variable: variable, data_pointer: data_pointer, data: field_data}
      end

      def replace_variable_in_placeholders_in_paragraphs(paragraphs, index, for_loop_placeholder_info, inbetween_placeholders)
        paragraphs.each do |paragraph|
          replace_variable_in_placeholders(index, for_loop_placeholder_info, inbetween_placeholders, paragraph)
        end
      end

    end
  end
end
