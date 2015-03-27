require 'active_support/core_ext/hash/indifferent_access'
require 'active_support/core_ext/object/blank'

module Word
  class PlaceholderEvaluator
    attr_accessor :placeholder
    def initialize(placeholder)
      self.placeholder = placeholder
    end

    def evaluate(data={})
      return "" if data.blank?
      placeholder_text = placeholder[:placeholder_text]
      field_identifier, options = placeholder_text[2..-3].split("|").map(&:strip)

      field_value = get_value_from_field_identifier(field_identifier, data)
      result = apply_options_to_field_value(field_value, options)
    end

    def get_value_from_field_identifier(field_identifier, data)
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

    def apply_options_to_field_value(field_value, options)
      field_value
    end
  end
end
