module Word
  module GroupOptions
    class FilterFields < Word::GroupOption
      #{{ loltest | list, show_lables: false, filter_fields: [a, b: [x y z], c, d, e, f, g] }}

      def apply_option
        if placeholder.field_value.present?
          placeholder.field_value = placeholder.field_value.map{|v| apply_field_filter_to_group_fields(v) }
        end
      end

      def apply_field_filter_to_group_fields(field_value)
        result = restrict_fields(parsed_field_filter, field_value)
      end

      def parsed_field_filter
        @allowed_fields ||= parse_nested_fields(params)
      end

      def parse_nested_fields(filter_part)
        #[a, b, c, d: [e,f: [x,y,z],g], x]
        #[a, b, c, d: [e,g,f: [x,y,z]], x]

        results = []
        ss = StringScanner.new(filter_part)
        loop do
          part = ss.scan_until(/[,\]]/)
          #part = ss.scan_until(/\]/)if part.blank?
          break if part.blank?

          if part.include?(':')
            whole_filter = part
            while(whole_filter.count('[') != whole_filter.count(']')) do
              whole_filter += ss.scan_until(/\]/)
            end
            group_name = whole_filter.split(':')[0].gsub(/[\[\],]/,'').strip
            fields = whole_filter.split(':')[1..-1].join(':').strip
            results << {group_name => parse_nested_fields(fields)}
            ss.scan_until(/,/) #Just to get to the end of the section
          else
            field = part.gsub(/[\[\],]/,'').strip
            results << field unless field.strip.blank?
            break if part.include?(']')
            # End when we get to ] in the normal string
          end

          break if ss.eos?
        end
        results
      end

      def restrict_fields(allowed_fields, fields)
        result = {}
        allowed_fields.each do |field|
          field_name = field.is_a?(String) ? field : field.keys.first
          next if fields[field_name].blank?

          if field.is_a?(String) || !Word::Placeholder.is_group_answer?(fields[field_name])
            result[field_name] = fields[field_name]
          else
            if fields[field_name].is_a? Array
              result[field_name] = []
              fields[field_name].each do |repeat_group|
                result[field_name] << restrict_fields(field[field_name], repeat_group)
              end
            else
              result[field_name] = restrict_fields(field[field_name], fields[field_name])
            end
          end
        end
        result
      end

    end
  end
end
