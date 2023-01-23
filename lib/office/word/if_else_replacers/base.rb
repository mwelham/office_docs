require 'liquid'

module Word
  module IfElseReplacers
    class Base
      attr_accessor :main_doc, :data, :options
      def initialize(main_doc, data, options = {})
        self.main_doc = main_doc
        self.data = data
        self.options = options
      end

      def evaluate_if(placeholder)
        expression = parse_if_else_placeholder(placeholder)[:expression]
        evaluate_expression(expression)
      end

      def parse_if_else_placeholder(placeholder)
        result = placeholder.gsub('{%','').gsub('%}','').match(/if (.+)/)
        expression = result[1].strip.gsub(/[“”]/, "\"").gsub(/[‘’]/, "\'") if !result.nil?

        raise "Invalid syntax for if placeholder #{placeholder}" if result.blank? || expression.blank?
        {expression: expression}
      end

      def evaluate_expression(expression)
        expression = expression.gsub(/\s=\s/," == ")
        expression = expression.gsub(/!(?<word>[\w\.]+)/,'\k<word> == null')
        expression = expression.gsub(/\sincludes\s/, ' contains ')

        liquid_if = ::Liquid::If.send(:new,'if',expression,{})
        condition = liquid_if.instance_variable_get("@blocks").first
        context_data = sanitize_data(data)
        context_data["yes"] = "yes"
        context_data["no"] = "no"
        context = Liquid::Context.new(context_data)
        result = condition.evaluate(context)
        result != false && result.present?
      end

      def sanitize_data(hash)
        sanitized_hash = {}
        hash.each do |k, v|
          if v.is_a?(Hash)
            sanitized_hash[k] = sanitize_data(v)
          elsif v.is_a?(Array)
            if v.length > 1
              sanitized_hash[k] = v.map{|value| sanitize_value(value)}
            else
              sanitized_hash[k] = sanitize_value(v.first)
            end
          else
            sanitized_hash[k] = sanitize_value(v)
          end
        end
        sanitized_hash
      end

      def sanitize_value(value)
        case
        when value.class.to_s == "Magick::Image"
          value.to_s
        when is_number?(value)
          value.to_f
        when value.is_a?(Hash)
          sanitize_data(value)
        else
          value.presence
        end
      end

      def is_number? string
        return false if string.blank?
        true if Float(string) rescue false
      end

    end
  end
end
