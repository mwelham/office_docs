require 'rmagick'
require 'office/word/placeholder_evaluator/option'
require 'office/word/placeholder_evaluator/group_option'

# require all the options
Dir[File.join(File.dirname(__FILE__) + '/options', "**/*.rb")].each do |f|
  require f
end

# require all the group options
Dir[File.join(File.dirname(__FILE__) + '/group_options', "**/*.rb")].each do |f|
  require f
end

module Word
  class Placeholder
    attr_accessor :field_identifier, :field_value, :field_options, :render_options, :final_value
    def initialize(field_identifier, field_value, options_string)
      self.field_identifier = field_identifier
      self.field_value = field_value
      self.field_options = parse_options(options_string)

      self.render_options = {}
      self.final_value = nil
    end

    def replacement
      if !self.final_value.nil?
        self.final_value
      else
        field_options.each do |o|
          o.apply_option
        end
        self.final_value = self.field_value
      end
    end

    def is_text_answer?(value = self.field_value)
      value.kind_of?(String) || value.kind_of?(Numeric) || value.kind_of?(Date) || value.kind_of?(Time) || value.kind_of?(DateTime)
    end

    def is_image_answer?(value = self.field_value)
      value.kind_of?(Magick::Image) || value.is_a?(Magick::ImageList)
    end

    def is_map_answer?(value = self.field_value)
      value.is_a?(Array) && value.first.present? && (value.first.is_a?(Magick::Image) || value.first.is_a?(Magick::ImageList))
    end

    def is_group_answer?(value = self.field_value)
      self.class.is_group_answer?(value)
    end

    def self.is_group_answer?(value)
      value.kind_of?(Array) && (value.empty? || value.first.kind_of?(Hash))
    end

    private

    def parse_options(options_in_string_format)
      whole_options = split_options_on_commas(options_in_string_format)
      option_objects = whole_options.map{|o| Word::Option.build_option_object(o, self)}
    end

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

  end
end
