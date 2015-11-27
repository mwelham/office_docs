require 'office/word/for_loop_expanders/base'

module Word
  module ForLoopExpanders
    class LoopTableRow < Word::ForLoopExpanders::Base

      # Get start and end paragraph in row

      def expand_loop(start_placeholder, end_placeholder, inbetween_placeholders)
        container = start_placeholder[:paragraph_object].document
        row = get_row(start_placeholder, end_placeholder)

        start_placeholder[:paragraph_object].replace_first_with_empty_runs(start_placeholder[:placeholder_text]).last
        end_placeholder[:paragraph_object].replace_first_with_empty_runs(end_placeholder[:placeholder_text]).first

        duplicate_row = generate_new_row(container, row)

        for_loop_placeholder_info = parse_for_loop_placeholder(start_placeholder[:placeholder_text])
        field_data = for_loop_placeholder_info[:data].presence || []
        if field_data.length == 0
          # do nothing - placeholders will get cleared
        else
          # Get paragraphs in existing row
          # Do row 0 replace
          paragraphs = get_paragraphs_from_row(container, row)
          replace_variable_in_placeholders_in_paragraphs(paragraphs, 0, for_loop_placeholder_info, inbetween_placeholders)

          last_row = row
          #
          field_data[1..-1].each_with_index do |data_set, i|
            new_row = duplicate_row.clone
            paragraphs = get_paragraphs_from_row(container, new_row)
            replace_variable_in_placeholders_in_paragraphs(paragraphs, i+1, for_loop_placeholder_info, inbetween_placeholders)
            last_row.add_next_sibling(new_row)
            last_row = new_row
          end
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

      def generate_new_row(container, existing_row)
        new_row = existing_row.clone
      end

      def get_paragraphs_from_row(container, row)
        container_node = container
        paragraphs = []
        row.xpath(".//w:p").each { |p| paragraphs << Office::Paragraph.new(p, container_node) }
        paragraphs
      end

    end
  end
end
