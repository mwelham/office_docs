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
        raise "Invalid syntax for if placeholder #{placeholder}" if result.blank? || expression.blank?
        {expression: expression}
      end

      def evaluate_expression(expression)
        split_expression = expression.split(' ')
        left_raw = split_expression[0].try(:strip)
        operator = split_expression[1].try(:strip)
        right_raw = split_expression[2..-1].try(:join, ' ')

        # Special case
        if left_raw[0] == '!'
          operator = '!'
          left_raw = left_raw[1..-1]
        end

        left = parse_input(left_raw)
        right = parse_input(right_raw)

        result = case operator
          when nil
            left.present?
          when '!'
            left.blank?
          when '=', '=='
            left == right
          when '!=', '<>'
            left != right
          when "includes"
            left.to_s.include? right.to_s
          else
            raise "Invalid if expression: #{expression}."
        end

        result
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
