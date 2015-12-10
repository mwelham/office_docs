require 'office/word/for_loop_expanders/base'

module Word
  module ForLoopExpanders
    class LoopOverParagraphs < Word::ForLoopExpanders::Base

      # break paragraphs on start + end
      # loop paragraphs inbetween

      def expand_loop(start_placeholder, end_placeholder, inbetween_placeholders)
        container = start_placeholder[:paragraph_object].document
        target_nodes = get_inbetween_nodes(start_placeholder, end_placeholder)

        #This is the 0 run of our loop.
        # Get a set of runs for the duplicate runs
        duplicate_nodes = generate_node_set(target_nodes)

        for_loop_placeholder_info = parse_for_loop_placeholder(start_placeholder[:placeholder_text])

        field_data = for_loop_placeholder_info[:data].presence || []
        if field_data.length == 0
          target_nodes.each do |node|
            Word::Template.remove_node(node)
          end
        else
          target_paragraphs = get_paragraphs_from_nodes(container, target_nodes)
          replace_variable_in_placeholders_in_paragraphs(target_paragraphs, 0, for_loop_placeholder_info, inbetween_placeholders)
          last_node = target_nodes.last

          field_data[1..-1].each_with_index do |data_set, i|
            new_node_set = generate_node_set(duplicate_nodes)
            new_paragraph_set = get_paragraphs_from_nodes(container, new_node_set)
            replace_variable_in_placeholders_in_paragraphs(new_paragraph_set, i+1, for_loop_placeholder_info, inbetween_placeholders)

            new_node_set.each do |node|
              last_node.add_next_sibling(node)
              last_node = node
            end

          end
        end
      end

      def get_inbetween_nodes(start_placeholder, end_placeholder)
        start_paragraph = start_placeholder[:paragraph_object]
        end_paragraph = end_placeholder[:paragraph_object]
        document = start_paragraph.document

        container = start_paragraph.node.parent
        start_node = start_paragraph.node
        end_node = end_paragraph.node

        starts_with = start_paragraph.plain_text.start_with?(start_placeholder[:placeholder_text])
        ends_with = end_paragraph.plain_text.end_with?(end_placeholder[:placeholder_text])

        start_run = start_paragraph.replace_first_with_empty_runs(start_placeholder[:placeholder_text]).last
        end_run = end_paragraph.replace_first_with_empty_runs(end_placeholder[:placeholder_text]).first

        if start_paragraph.plain_text.gsub(" ", "").length == 0
          start_placeholder_paragraph = start_paragraph
          index = container.children.index(start_node)
          start_node = container.children[index + 1]
          document.remove_paragraph(start_placeholder_paragraph)
        else
          start_paragraph = starts_with ? start_paragraph : start_paragraph.split_after_run(start_run)
          start_node = start_paragraph.node
        end

        if end_paragraph.plain_text.gsub(" ", "").length == 0
          end_placeholder_paragraph = end_paragraph
          index = container.children.index(end_node)
          end_node = container.children[index - 1]
          document.remove_paragraph(end_placeholder_paragraph)
        else
          end_paragraph.split_after_run(end_run) if(!ends_with)
          end_node = end_paragraph.node
        end

        start_index = container.children.index(start_node)
        end_index = container.children.index(end_node)

        container.children[start_index..end_index]
      end

      def generate_node_set(nodes)
        nodes.map(&:clone)
      end

      def get_paragraphs_from_nodes(container, nodes)
        paragraphs = []
        nodes.each do |node|
          if node.name == 'p'
            paragraphs << Office::Paragraph.new(node, container)
          else
            node.xpath(".//w:p").each { |p| paragraphs << Office::Paragraph.new(p, container) }
          end
        end
        paragraphs
      end

    end
  end
end
