module Word
  class GroupPlaceholder < Word::Placeholder
    include Word::ImageFunctions

    attr_accessor :group_generation_options, :form_xml_def

    def initialize(field_identifier, field_value, options_string, form_xml_def=nil)
      super(field_identifier, field_value, options_string)
      self.form_xml_def = form_xml_def
      self.group_generation_options = {generation_method: :table, show_labels: true}

      # Changing hash to array for group - we turn single answer groups into a hash for accessing so need to change back here if they want the whole group printed.
      self.field_value = [field_value] if field_value.is_a? Hash
    end

    def replacement
      if !self.final_value.nil?
        self.final_value
      else
        field_options.sort_by(&:importance).reverse.each do |o|
          begin
            o.apply_option
          rescue => e
            raise "Error applying option #{o.class.to_s.underscore.humanize} to field #{self.field_identifier} - #{e.inspect}"
          end
        end
        self.final_value = if group_generation_options[:generation_method] == :list
          create_list_for_group(form_xml_def, field_identifier.gsub('fields.',''), field_value, group_generation_options)
        else
          create_table_for_group(form_xml_def, field_identifier.gsub('fields.',''), field_value, group_generation_options)
        end
      end
    end

    private

    def parse_options(options_in_string_format)
      whole_options = split_options_on_commas(options_in_string_format)
      option_objects = whole_options.map{|o| Word::GroupOption.build_option_object(o, self)}.compact
    end

    def create_list_for_group(form_xml_def, group_id, values, options = {}, indent = "")
        return "" if values.blank?
        list = []
        values.each_index do |i|
          values[i].each do |id, answer|
            next if answer.blank?
            full_id = "#{group_id}.#{id}"
            title = form_xml_def.blank? ? full_id : form_xml_def.get_field_label(full_id)
            should_show_labels = options[:show_labels].nil? || options[:show_labels] != false

            case
              when is_text_answer?(answer)
                add_to_list(list, title, answer, indent, {on_same_line: true, should_show_labels: should_show_labels})
              when is_image_answer?(answer)
                answer = resize_image_answer(answer, options[:image_size][:width], options[:image_size][:height]) if options[:image_size].present?
                add_to_list(list, title, answer, indent, {on_same_line: false, should_show_labels: should_show_labels})
              when is_map_answer?(answer)
                answer[0] = resize_image_answer(answer[0], options[:map_size][:width], options[:map_size][:height]) if options[:map_size].present?
                add_to_list(list, title, answer, indent, {on_same_line: false, should_show_labels: should_show_labels})
              when is_group_answer?(answer)
                answer_to_add = create_list_for_group(form_xml_def, full_id, answer, options, "-\t\t\t#{indent}")
                add_to_list(list, title, answer_to_add, indent, {on_same_line: false, should_show_labels: should_show_labels})
              else
                add_to_list(list, title, answer, indent, {on_same_line: false, should_show_labels: should_show_labels})
            end
          end
          list << "" unless i == values.length - 1
        end
        list
      end

      def add_to_list(list, title, answer, indent, options = {})
        on_same_line = options[:on_same_line]
        should_show_labels = options[:should_show_labels]

        on_same_line = false if on_same_line.nil?
        should_show_labels = true if should_show_labels.nil?

        if should_show_labels
          if on_same_line
            list << "#{indent}#{title}\t\t\t#{answer}"
          else
            list << "#{indent}#{title}" << answer
          end
        else
          if on_same_line
            list << "#{indent}#{answer}"
          else
            list << answer
          end
        end
      end

      def create_table_for_group(form_xml_def, group_id, values, options = {})
        return "" if values.blank?
        table = {}
        headers = {}
        # Set up headers
        values.collect{|v| v.keys}.flatten.uniq.each do |answer_key|
          full_id = "#{group_id}.#{answer_key}"
          title = form_xml_def.blank? ? full_id : form_xml_def.get_field_label(full_id)
          headers[answer_key] = {title: title, full_id: full_id}
          table[title] ||= []
        end

        # Fill in values
        values.each do |answer_set|
          headers.each do |key, attributes|
            title = attributes[:title]
            full_id = attributes[:full_id]
            answer = answer_set[key] || ""

            case
              when is_group_answer?(answer)
                table[title] << create_table_for_group(form_xml_def, full_id, answer, options)
              when is_image_answer?(answer)
                width = 500
                height = 500
                if options[:image_size].present?
                  width = options[:image_size][:width]
                  height = options[:image_size][:height]
                end
                table[title] << resize_image_answer(answer, width, height)
              when is_map_answer?(answer)
                answer[0] = resize_image_answer(answer[0], options[:map_size][:width], options[:map_size][:height]) if options[:map_size].present?
                table[title] << answer
              else
                table[title] << answer
            end
          end
        end
        table
      end

  end
end
