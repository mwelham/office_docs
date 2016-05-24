require 'office/word/for_loop_expanders/base'

module Word
  module IfElseReplacers
    class IfElseTableRow < Word::ForLoopExpanders::Base

      def replace_if_else(start_placeholder, end_placeholder, inbetween_placeholders)
        container = start_placeholder[:paragraph_object].document
        row = get_row(start_placeholder, end_placeholder)
        should_keep = evaluate_if(start_placeholder[:placeholder_text])
        if should_keep
          start_placeholder[:paragraph_object].replace_first_with_empty_runs(start_placeholder[:placeholder_text]).last
          end_placeholder[:paragraph_object].replace_first_with_empty_runs(end_placeholder[:placeholder_text]).first
        else
          Word::Template.remove_node(row)
        end
      end

      def get_row(start_placeholder, end_placeholder)
        row = nil
        current = start_placeholder[:paragraph_object].node
        while row == nil
          parent = current.parent
          raise "No row object ???" if parent == nil
          if parent.name == 'tr'
            row = parent
          else
            current = current.parent
          end
        end
        row
      end

    end
  end
end
