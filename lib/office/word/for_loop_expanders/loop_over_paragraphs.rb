require 'office/word/for_loop_expanders/base'

module Word
  module ForLoopExpanders
    class LoopOverParagraphs < Word::ForLoopExpanders::Base

      # break paragraphs on start + end
      # loop paragraphs inbetween

      def expand_loop(start_placeholder, end_placeholder, inbetween_placeholders)
        target_paragraphs = get_paragraphs(start_placeholder, end_placeholder)

        container = target_paragraphs.first.document

        #This is the 0 run of our loop.
        # Get a set of runs for the duplicate runs
        duplicate_paragraphs = generate_new_paragraph_set(container, target_paragraphs)
        for_loop_placeholder_info = parse_for_loop_placeholder(start_placeholder[:placeholder_text])

        field_data = for_loop_placeholder_info[:data].presence || []
        if field_data.length == 0
          target_paragraphs.each do |paragraph|
            container.remove_paragraph(paragraph)
          end
        else
          replace_variable_in_placeholders_in_paragraphs(target_paragraphs, 0, for_loop_placeholder_info, inbetween_placeholders)
          last_paragraph = target_paragraphs.last
          field_data[1..-1].each_with_index do |data_set, i|
            new_paragraph_set = generate_new_paragraph_set(container, duplicate_paragraphs)
            replace_variable_in_placeholders_in_paragraphs(new_paragraph_set, i+1, for_loop_placeholder_info, inbetween_placeholders)

            new_paragraph_set.each do |paragraph|
              container.insert_new_paragraph_object_after_paragraph(last_paragraph, paragraph)
              last_paragraph = paragraph
            end
            #last_paragraph = new_paragraph_set.last
          end
        end

      end

      def get_paragraphs(start_placeholder, end_placeholder)
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

        start_paragraph_index = document.paragraphs.index(start_paragraph)
        end_paragraph_index = document.paragraphs.index(end_paragraph)
        document.paragraphs[start_paragraph_index..end_paragraph_index]
      end

      def generate_new_paragraph_set(container, paragraphs)
        new_paragraph_set = []

        paragraphs.each do |p|
          new_p = Office::Paragraph.new(p.node.clone, container)
          new_paragraph_set << new_p
        end

        new_paragraph_set
      end

      def replace_variable_in_placeholders_in_paragraphs(paragraphs, index, for_loop_placeholder_info, inbetween_placeholders)
        paragraphs.each do |paragraph|
          replace_variable_in_placeholders(index, for_loop_placeholder_info, inbetween_placeholders, paragraph)
        end
      end


    end
  end
end
