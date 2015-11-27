require 'office/word/for_loop_expanders/loop_in_paragraph'
require 'office/word/for_loop_expanders/loop_over_paragraphs'
require 'office/word/for_loop_expanders/loop_table_row'

module Word
  class ForLoopExpander
    attr_accessor :main_doc, :data, :options, :placeholders
    def initialize(main_doc, data, options = {})
      self.main_doc = main_doc
      self.data = data
      self.options = options
    end

    def expand_for_loops(container)
      # Get placeholders in paragraphs
      paragraphs = container.paragraphs
      self.placeholders = Word::PlaceholderFinder.get_placeholders(paragraphs)
      while there_are_for_loop_placeholders(placeholders)
        i = 0
        while i < placeholders.length
          start_placeholder = placeholders[i]
          if start_placeholder[:placeholder_text].include?("foreach")
            end_index = get_end_index(i)
            expand_loop(i, end_index)

            i = end_index + 1
          else
            i += 1
          end
        end
        paragraphs = resync_container(container)
        self.placeholders = Word::PlaceholderFinder.get_placeholders(paragraphs)
      end

      # i = 0
      # While we have placeholders
        # Get first for start
        # Loop through placeholders to get its end
        # Expand the for loop out replacing the variable placeholders
        # expand the for loops in each set generated
        # i = index_of(end_placeholder)

    end

    def get_end_index(start_index)
      level = 0
      placeholders[(start_index+1)..-1].each_with_index do |p, j|
        if p[:placeholder_text].include?("endeach") && level == 0
          return (start_index+1)+j
        elsif p[:placeholder_text].include?("endeach") && level > 0
          level -= 1
        elsif p[:placeholder_text].include?("foreach")
          level += 1
        end
      end
      raise "For each missing an end"
    end

    def expand_loop(start_index, end_index)
      start_placeholder = placeholders[start_index]
      end_placeholder = placeholders[end_index]
      inbetween_placeholders = placeholders[(start_index+1)..(end_index-1)]
      #puts "Expanding from\n #{start_placeholder.inspect}\nto\n#{end_placeholder.inspect}\n\n"
      if start_placeholder[:paragraph_index] == end_placeholder[:paragraph_index]
        # if start and end are in the same paragraph
        looper = Word::ForLoopExpanders::LoopInParagraph.new(main_doc, data, options)
        looper.expand_loop(start_placeholder, end_placeholder, inbetween_placeholders)
      elsif placeholders_are_in_different_table_cells_in_same_row?(start_placeholder, end_placeholder)
        # else if start is in a table cell and end is in a different table cell
        # loop whole row
        looper = Word::ForLoopExpanders::LoopTableRow.new(main_doc, data, options)
        looper.expand_loop(start_placeholder, end_placeholder, inbetween_placeholders)
      elsif start_placeholders_is_in_table_cell_but_end_is_not_in_row?(start_placeholder, end_placeholder)
        # else if start is in a table cell but end is not in a table cell at all
        # raise error
        raise "For loop start and end mismatch - start is in table row but no end"
      else
        # else its over paragraphs
        looper = Word::ForLoopExpanders::LoopOverParagraphs.new(main_doc, data, options)
        looper.expand_loop(start_placeholder, end_placeholder, inbetween_placeholders)
      end
    end

    def placeholders_are_in_different_table_cells_in_same_row?(start_placeholder, end_placeholder)
      start_placeholder_parent = start_placeholder[:paragraph_object].node.parent
      end_placeholder_parent = end_placeholder[:paragraph_object].node.parent

      start_placeholder_parent.name == 'tc' &&
      end_placeholder_parent.name == 'tc' &&
      start_placeholder_parent != end_placeholder_parent &&
      start_placeholder_parent.parent == end_placeholder_parent.parent
    end

    def start_placeholders_is_in_table_cell_but_end_is_not_in_row?(start_placeholder, end_placeholder)
      start_placeholder_parent = start_placeholder[:paragraph_object].node.parent
      end_placeholder_parent = end_placeholder[:paragraph_object].node.parent

      return false if start_placeholder_parent.name != 'tc'

      # First one is in a table cell - just need to make sure end is a tc and in the same row
      end_placeholder_parent.name != 'tc' ||
      start_placeholder_parent.parent != end_placeholder_parent.parent
    end

    def resync_container(container)
      container.parse_paragraphs(container.container_node)
      paragraphs = container.paragraphs
    end

    def there_are_for_loop_placeholders(placeholders)
      placeholders.any?{|p| p[:placeholder_text].include?("foreach") }
    end

  end#endclass
end#endmodule


__END__

{
  placeholder_text: '{{ i_am_holder }}',
  paragraph_index: 0,
  begin: {run_index: 3, char_index: 15},
  end: {run_index: 5, char_index: 2}
}


To add a node somewhere

preceding_r_node = @node.add_child(@node.document.create_element("r"))
populate_r_node(preceding_r_node, text)
first_run.node.add_previous_sibling(preceding_r_node)
@runs.insert(@runs.index(first_run), Run.new(preceding_r_node, self))
