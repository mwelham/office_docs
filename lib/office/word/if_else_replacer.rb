require 'office/word/if_else_replacers/if_else_in_paragraph'
require 'office/word/if_else_replacers/if_else_over_paragraphs'
require 'office/word/if_else_replacers/if_else_table_row'

require 'office/word/placeholder_position_check_methods'

module Word
  class IfElseReplacer
    include PlaceholderPositionCheckMethods

    attr_accessor :main_doc, :data, :options, :placeholders

    IF_ELSE_START_MATCHER = /\W(if)/
    IF_ELSE_END_MATCHER = /endif/

    def initialize(main_doc, data, options = {})
      self.main_doc = main_doc
      self.data = data
      self.options = options
    end

    # Have to resync after each if/else replacement because of the way the if/else in paragraph works.
    #TODO - instead of resyncing the whole container just resync the paragraphs and placeholders affected.
    # Do it in replace_if_else - get the start and end paragraphs, uniq them
    # Get the placeholders in those paragraphs - delete them
    # Reget the placeholders in those paragraphs
    # Add them to the placeholders list
    # Sort the list

    def replace_all_if_else(container)
      # Get placeholders in paragraphs
      paragraphs = container.paragraphs
      self.placeholders = Word::PlaceholderFinder.get_placeholders(paragraphs)
      while there_are_if_else_placeholders?(placeholders)
        i = 0
        while i < placeholders.length
          start_placeholder = placeholders[i]
          if start_placeholder[:placeholder_text].match(IF_ELSE_START_MATCHER)
            end_index = get_end_index(i)
            raise "Missing endif for if placeholder: #{start_placeholder[:placeholder_text]}" if end_index.nil?
            replace_if_else(i, end_index)
            paragraphs = resync_container(container)
            self.placeholders = Word::PlaceholderFinder.get_placeholders(paragraphs)
            break
          else
            i += 1
          end
        end
      end

    end

    def get_end_index(start_index)
      level = 0
      placeholders[(start_index+1)..-1].each_with_index do |p, j|
        if p[:placeholder_text].match(IF_ELSE_END_MATCHER) && level == 0
          return (start_index+1)+j
        elsif p[:placeholder_text].match(IF_ELSE_END_MATCHER) && level > 0
          level -= 1
        elsif p[:placeholder_text].match(IF_ELSE_START_MATCHER)
          level += 1
        end
      end
      nil
    end

    def replace_if_else(start_index, end_index)
      start_placeholder = placeholders[start_index]
      end_placeholder = placeholders[end_index]
      inbetween_placeholders = placeholders[(start_index+1)..(end_index-1)]

      if start_placeholder[:paragraph_index] == end_placeholder[:paragraph_index]
        # if start and end are in the same paragraph
        looper = Word::IfElseReplacers::IfElseInParagraph.new(main_doc, data, options)
        looper.replace_if_else(start_placeholder, end_placeholder, inbetween_placeholders)
      elsif placeholders_are_in_different_table_cells_in_same_row?(start_placeholder, end_placeholder)
        looper = Word::IfElseReplacers::IfElseTableRow.new(main_doc, data, options)
        looper.replace_if_else(start_placeholder, end_placeholder, inbetween_placeholders)
      elsif start_placeholders_is_in_table_cell_but_end_is_not_in_row?(start_placeholder, end_placeholder)
        # else if start is in a table cell but end is not in a table cell at all
        # raise error
        raise "If statement start and end mismatch - start is in table row but no end: #{start_placeholder[:placeholder_text]}"
      elsif placeholders_are_in_different_containers?(start_placeholder, end_placeholder)
        # else if start is in a table cell but end is not in a table cell at all
        # raise error
        raise "If start and end are in different containers for if #{start_placeholder[:placeholder_text]}"
      else
        # else its over paragraphs
        looper = Word::IfElseReplacers::IfElseOverParagraphs.new(main_doc, data, options)
        looper.replace_if_else(start_placeholder, end_placeholder, inbetween_placeholders)
      end
    end

    def there_are_if_else_placeholders?(placeholders)
      placeholders.any?{|p| p[:placeholder_text].match(IF_ELSE_START_MATCHER) }
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
