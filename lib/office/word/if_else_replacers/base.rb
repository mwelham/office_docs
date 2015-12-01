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
        expression = result[1].strip.gsub(/[“”]/, "\"") if !result.nil?
        raise "Invalid syntax for foreach placeholder #{placeholder}" if result.blank? || expression.blank?
        {expression: expression}
      end

      def evaluate_expression(expression)
        left, operator, right = expression.split(' ')
        left = parse_input(left)
        right = parse_input(right)
        case operator
          when nil
            left.present?
          when '=', '=='
            left == right
          when '!=', '<>'
            left != right
          else
            raise "Invalid if expression: #{expression}."
        end
      end

      def parse_input(input)
        return nil if input.nil?
        case
        when is_number?(input)
          Float(input)
        when is_data?(input)
          Word::Template.get_value_from_field_identifier(input, data)
        else
          input.strip.gsub("\"","")
        end
      end

      def is_number? string
        true if Float(string) rescue false
      end

      def is_data? string
        !string.include?("\"") &&
        data.keys.include?(string.split('.').first)
      end

    end
  end
end
