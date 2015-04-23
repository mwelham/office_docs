require 'active_support/core_ext/hash/indifferent_access'
require 'active_support/core_ext/object/blank'
require 'action_view'

include ActionView::Helpers::NumberHelper

module Word
  class PlaceholderEvaluator
    attr_accessor :placeholder    #Placeholder to work out
    attr_accessor :replacement    #Result from the placeholder
    attr_accessor :render_options #Options to be passed on to the renderer

    def initialize(placeholder)
      self.placeholder = placeholder
      self.render_options = {}
    end

    def evaluate(data={}, global_options={})
      return "" if data.blank?
      placeholder_text = placeholder[:placeholder_text]
      field_identifier, field_options = placeholder_text[2..-3].split("|").map(&:strip) #2,-3 to get rid of {{ and }}
      field_options = split_options_on_commas(field_options)

      field_value = get_value_from_field_identifier(field_identifier, data, global_options)
      self.replacement = apply_options_to_field_value(field_identifier, field_value, field_options, global_options)
    end

    def get_value_from_field_identifier(field_identifier, data, options={})
      result = data.with_indifferent_access
      field_recurse = field_identifier.split('.')
      field_recurse.each do |identifier|
        if result.is_a? Array
          result = result.length == 1 ? result.first : {}
        end

        result = result[identifier]
        break if result == nil
      end
      result
    end

    def apply_options_to_field_value(field_identifier, field_value, field_options, global_options)
      field_options ||= []

      field_value = case
      when is_text_answer?(field_value)
        apply_options_to_normal_value(field_value, field_options)
      when (field_value.is_a?(Magick::Image) || field_value.is_a?(Magick::ImageList))
        apply_options_to_image_value(field_value, field_options)
      when is_map_answer?(field_value)
        apply_options_to_map_value(field_value, field_options)
      else #must be a group
        apply_options_to_group_value(field_identifier, field_value, field_options, global_options)
      end

      field_value
    end

    def apply_options_to_normal_value(field_value, field_options)
      field_options.each do |option|
        case option.downcase
        when "currency"
          field_value = number_to_currency(field_value.to_f, unit: '')
        when "each_answer_on_new_line"
          field_value = field_value.split(',').map(&:strip)
        else
          field_value.to_s
        end
      end
      field_value
    end

    def apply_options_to_image_value(field_value, field_options)
      field_options.each do |option|
        edges = /(\d+)x(\d+)/.match(option)
        field_value = resize_image_answer(field_value, edges[1].to_f, edges[2].to_f) if edges.present? && edges[1].present? && edges[2].present?
      end
      field_value
    end

    def apply_options_to_map_value(field_value, field_options)
      # Image size options
      field_value[0] = apply_options_to_image_value(field_value[0], field_options)

      # Hyperlink options
      hyperlink_option = get_option_from_field_options(field_options, 'hyperlink')
      if hyperlink_option[:params].downcase == 'true'
        coord_info = field_value[1]
        lat = coord_info.match(/lat=((\d+|-\d)+\.\d+)/)
        long = coord_info.match(/long=((\d+|-\d)+\.\d+)/)
        if !lat.nil? && !long.nil?
          render_options[:hyperlink] = "http://maps.google.com/?q=#{lat[1]},#{long[1]}"
        end
      end

      # Coordinate info options
      coordinate_info_option = get_option_from_field_options(field_options, 'show_coordinate_info')
      if coordinate_info_option[:params].downcase == 'false'
        field_value = [field_value[0]]
      end

      field_value
    end

    def apply_options_to_group_value(field_identifier, field_value, field_options, global_options={})
      form_xml_def = global_options[:form_xml_def]

      # Changing hash to array for group - we turn single answer groups into a hash for accessing so need to
      # change back here if they want the whole group printed.
      field_value = Array(field_value) if field_value.is_a? Hash

      #apply field filter if it is present
      field_filter = field_options.select{|o| o.downcase.include?('filter_fields')}.first
      field_value = field_value.map{|v| apply_field_filter_to_group_fields(field_filter, v) } if field_filter.present?

      #Image/Map size settings
      creation_options = {}
      ['image_size', 'map_size'].each do |size_setting|
        size = field_options.select{|o| o.downcase.include?(size_setting)}.first
        if size.present?
          edges = /(\d+)x(\d+)/.match(size)
          creation_options[size_setting.to_sym] = {}
          creation_options[size_setting.to_sym][:width] = edges[1].to_f
          creation_options[size_setting.to_sym][:height] = edges[2].to_f
        end
      end

      if field_options.any?{|o| o.downcase == 'list'}
        field_value = create_list_for_group(form_xml_def, field_identifier.gsub('fields.',''), field_value, creation_options)
      else
        global_options[:use_full_width] = true
        field_value = create_table_for_group(form_xml_def, field_identifier.gsub('fields.',''), field_value, creation_options)
      end
      field_value
    end

    #{{ loltest | list, filter_fields: [a, b: [x y z], c, d, e, f, g] }}

    #
    #
    #####
    ##### Helper functions for applying options under here
    #####
    #
    #

    #
    #
    ## Image Functions
    #
    #

    def resize_image_answer(image, width, height)
      return image if image.nil? or image == "" or width < 1 or height < 1
      return image if image.columns < width and image.rows < height
      image.resize([1.0 * width / image.columns, 1.0 * height / image.rows].min)
    end

    #
    #
    ## Group Functions
    #
    #

    def apply_field_filter_to_group_fields(field_filter, field_value)
      parsed_field_filter = parse_field_filter(field_filter)
      result = restrict_fields(parsed_field_filter, field_value)
    end

    def parse_field_filter(field_filter)
      # Get rid of "fields:"
      fields = field_filter.split(':')[1..-1].join(':').strip
      allowed_fields = parse_nested_fields(fields)
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

        if field.is_a?(String) || !is_group_answer?(fields[field_name])
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

    def create_list_for_group(form_xml_def, group_id, values, options = {}, indent = "")
      return "" if values.blank?
      list = []
      values.each_index do |i|
        values[i].each do |id, answer|
          next if answer.blank?
          full_id = "#{group_id}.#{id}"
          title = form_xml_def.blank? ? full_id : form_xml_def.get_field_label(full_id)
          case
            when is_text_answer?(answer)
              list << "#{indent}#{title}\t\t\t#{answer}"
            when is_image_answer?(answer)
              answer = resize_image_answer(answer, options[:image_size][:width], options[:image_size][:height]) if options[:image_size].present?
              list << "#{indent}#{title}" << answer
            when is_map_answer?(answer)
              answer[0] = resize_image_answer(answer[0], options[:map_size][:width], options[:map_size][:height]) if options[:map_size].present?
              list << "#{indent}#{title}" << answer
            when is_group_answer?(answer)
              list << "#{indent}#{title}" << create_list_for_group(form_xml_def, full_id, answer, options, "-\t\t\t#{indent}")
            else
              list << "#{indent}#{title}" << answer
          end
        end
        list << "" unless i == values.length - 1
      end
      list
    end

    def create_table_for_group(form_xml_def, group_id, values, options = {})
      return "" if values.blank?
      table = {}
      values.each do |v|
        v.each do |id, answer|
          full_id = "#{group_id}.#{id}"
          title = form_xml_def.blank? ? full_id : form_xml_def.get_field_label(full_id)
          table[title] = [] unless table.has_key?(title)
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

    #
    #
    ## MISC FUNCTIONS
    #
    #

    def split_options_on_commas(options_in_string_format)
      results = []
      return results if options_in_string_format.blank?

      ss = StringScanner.new(options_in_string_format)
      loop do
        part = ss.scan_until(/[,\]]/)
        (part = ss.rest) and ss.terminate if part.blank? && !ss.eos?

        break if part.blank?

        if part.include?(':') && part.include?('[')
          whole_filter = part
          while(whole_filter.count('[') != whole_filter.count(']')) do
            next_part = ss.scan_until(/\]/)
            raise "Missing ] while parsing options." if next_part.blank?
            whole_filter += next_part
          end
          ss.scan_until(/,/) #Just to get to the end of the section
          results << whole_filter.strip
        else
          option = part.gsub(/[\[\],]/,'').strip
          results << option unless option.strip.blank?
          break if part.include?(']')
          # End when we get to ] in the normal string
        end

        break if ss.eos?
      end
      results
    end

    def is_group_answer?(value)
      value.kind_of?(Array) && (value.empty? || value.first.kind_of?(Hash))
    end

    def is_text_answer?(value)
      value.kind_of?(String) || value.kind_of?(Numeric) || value.kind_of?(Date) || value.kind_of?(Time)
    end

    def is_image_answer?(value)
      value.kind_of?(Magick::Image)
    end

    def is_map_answer?(value)
      value.is_a?(Array) && value.first.present? && value.first.is_a?(Magick::Image)
    end

    def get_option_from_field_options(field_options, option)
      option_text = field_options.select{|o| o.match(/#{option}\s*:?/)}.first
      option, params = option_text.split(':').map(&:strip)
      {option: option, params: params}
    end

  end#endclass
end
