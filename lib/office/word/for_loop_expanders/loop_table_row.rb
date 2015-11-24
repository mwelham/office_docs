require 'office/word/for_loop_expanders/base'

module Word
  module ForLoopExpanders
    class LoopTableRow < Word::ForLoopExpanders::Base

      # Get start and end paragraph in row

      def expand_loop(start_placeholder, end_placeholder, inbetween_placeholders)
        container = start_placeholder[:paragraph_object].document
        row = get_row(start_placeholder, end_placeholder)

        duplicate_row = generate_new_row(container, row_object)

        field_data = for_loop_placeholder_info[:data].presence || []
        if field_data.length == 0
          # do nothing - placeholders will get cleared
        else
          # Get paragraphs in existing row
          # Do row 0 replace

          #
          field_data[1..-1].each_with_index do |data_set, i|
            # Add duplicate row object after last row
            # Get paragraphs in duplicate row
            # Generate and insert new paragraph objects after last one in existing row
            # Do row i replace
          end
        end

      end

      def get_row(start_placeholder, end_placeholder)
        row = nil
        current = start_placeholder[:paragraph_object]
        while row == nil
          parent = current.parent
          raise "No row object ???" if parent == nil
          row = parent if parent.name == 'tr'
        end
        row
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
