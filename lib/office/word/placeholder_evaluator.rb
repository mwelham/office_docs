require 'active_support/core_ext/hash/indifferent_access'
require 'active_support/core_ext/object/blank'
require 'action_view'

require 'office/word/placeholder_evaluator/shared/image_functions'

require 'office/word/placeholder_evaluator/placeholder'
require 'office/word/placeholder_evaluator/group_placeholder'

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

    #Global options is a bit of a hack - its stuff passed in from the outside world, like xml_def
    def evaluate(data={}, global_options={})
      return {replacement: "", render_options: {}} if data.blank?
      placeholder_text = placeholder[:placeholder_text]
      field_identifier, field_options = placeholder_text[2..-3].split("|").map(&:strip) #2,-3 to get rid of {{ and }}

      field_value = get_value_from_field_identifier(field_identifier, data, global_options)

      placeholder = if is_group_answer?(field_value)
        Word::GroupPlaceholder.new(field_identifier, field_value, field_options, global_options[:form_xml_def])
      else
        Word::Placeholder.new(field_identifier, field_value, field_options)
      end
      replacement = placeholder.replacement
      render_options = placeholder.render_options

      return {replacement: replacement, render_options: render_options}
    end

    def get_value_from_field_identifier(field_identifier, data, options={})
      result = data.with_indifferent_access
      field_recurse = field_identifier.split('.')
      field_recurse.each do |identifier|
        if result.is_a? Array
          result = result.length == 1 ? result.first : {}
        end

        result = result[identifier]
        if result == nil
          result = ""
          break
        end
      end
      result
    end


    def is_group_answer?(value)
      value.kind_of?(Array) && (value.empty? || value.first.kind_of?(Hash))
    end

  end#endclass
end
