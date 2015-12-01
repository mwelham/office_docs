require 'office/word/if_else_replacers/base'

module Word
  module IfElseReplacers
    class IfElseOverParagraphs < Word::IfElseReplacers::Base

      def replace_if_else(start_placeholder, end_placeholder, inbetween_placeholders)
        container = start_placeholder[:paragraph_object].document
        target_nodes = get_inbetween_nodes(start_placeholder, end_placeholder)

        should_keep = evaluate_if(start_placeholder[:placeholder_text])

        if !should_keep
          target_nodes.each(&:remove)
        end
      end

      def get_inbetween_nodes(start_placeholder, end_placeholder)
        start_paragraph = start_placeholder[:paragraph_object]
        end_paragraph = end_placeholder[:paragraph_object]
        document = start_paragraph.document

        starts_with = start_paragraph.plain_text.start_with?(start_placeholder[:placeholder_text])
        ends_with = end_paragraph.plain_text.end_with?(end_placeholder[:placeholder_text])

        start_run = start_paragraph.replace_first_with_empty_runs(start_placeholder[:placeholder_text]).last
        end_run = end_paragraph.replace_first_with_empty_runs(end_placeholder[:placeholder_text]).first

        if start_paragraph.plain_text.gsub(" ", "").length == 0
          start_placeholder_paragraph = start_paragraph
          index = document.paragraphs.index(start_placeholder_paragraph)
          start_paragraph = document.paragraphs[(index+1)]
          document.remove_paragraph(start_placeholder_paragraph)
        else
          start_paragraph = starts_with ? start_paragraph : start_paragraph.split_after_run(start_run)
        end

        if end_paragraph.plain_text.gsub(" ", "").length == 0
          end_placeholder_paragraph = end_paragraph
          index = document.paragraphs.index(end_paragraph)
          end_paragraph = document.paragraphs[(index-1)]
          document.remove_paragraph(end_placeholder_paragraph)
        else
          end_paragraph.split_after_run(end_run) if(!ends_with)
        end

        container = start_paragraph.node.parent
        start_index = container.children.index(start_paragraph.node)
        end_index = container.children.index(end_paragraph.node)
        container.children[start_index..end_index]
      end

    end
  end
end
