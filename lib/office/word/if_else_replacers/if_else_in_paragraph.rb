require 'office/word/if_else_replacers/base'

module Word
  module IfElseReplacers
    class IfElseInParagraph < Word::IfElseReplacers::Base

      # Break runs on start + end
      # loop runs inbetween

      def replace_if_else(start_placeholder, end_placeholder, inbetween_placeholders)
        paragraph = start_placeholder[:paragraph_object]

        start_run = paragraph.replace_first_with_empty_runs(start_placeholder[:placeholder_text]).last
        end_run = paragraph.replace_first_with_empty_runs(end_placeholder[:placeholder_text]).first

        from_run = paragraph.runs.index(start_run) + 1
        to_run = paragraph.runs.index(end_run) - 1

        inbetween_runs = paragraph.runs[from_run..to_run]
        should_keep = evaluate_if(start_placeholder[:placeholder_text])

        if !should_keep
          inbetween_runs.each do |run|
            paragraph.remove_run(run)
          end
          if paragraph.plain_text.gsub(/\s/, "").length == 0
            Word::Template.remove_node(paragraph.node)
          end
        end

      end
    end
  end
end
