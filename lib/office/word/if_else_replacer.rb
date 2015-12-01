require 'office/word/if_else_replacers/if_else_in_paragraph'
require 'office/word/if_else_replacers/if_else_over_paragraphs'

module Word
  class IfElseReplacer
    attr_accessor :main_doc, :data, :options, :placeholders
    def initialize(main_doc, data, options = {})
      self.main_doc = main_doc
      self.data = data
      self.options = options
    end

    def replace_all_if_else(container)
      # Get placeholders in paragraphs
      paragraphs = container.paragraphs
      self.placeholders = Word::PlaceholderFinder.get_placeholders(paragraphs)
      while there_are_if_else_placeholders?(placeholders)
        i = 0
        while i < placeholders.length
          start_placeholder = placeholders[i]
          if start_placeholder[:placeholder_text].include?("if ")
            end_index = get_end_index(i)
            replace_if_else(i, end_index)

            i = end_index + 1
          else
            i += 1
          end
        end
        paragraphs = resync_container(container)
        self.placeholders = Word::PlaceholderFinder.get_placeholders(paragraphs)
      end

    end

    def get_end_index(start_index)
      level = 0
      placeholders[(start_index+1)..-1].each_with_index do |p, j|
        if p[:placeholder_text].include?("endif") && level == 0
          return (start_index+1)+j
        elsif p[:placeholder_text].include?("endif") && level > 0
          level -= 1
        elsif p[:placeholder_text].include?("if ")
          level += 1
        end
      end
      raise "if statement missing an end"
    end

    def replace_if_else(start_index, end_index)
      start_placeholder = placeholders[start_index]
      end_placeholder = placeholders[end_index]
      inbetween_placeholders = placeholders[(start_index+1)..(end_index-1)]
      if start_placeholder[:paragraph_index] == end_placeholder[:paragraph_index]
        # if start and end are in the same paragraph
        looper = Word::IfElseReplacers::IfElseInParagraph.new(main_doc, data, options)
        looper.replace_if_else(start_placeholder, end_placeholder, inbetween_placeholders)
      elsif if_else_are_in_different_container?(start_placeholder, end_placeholder)
        # else if start is in a table cell but end is not in a table cell at all
        # raise error
        raise "If start and end are in different containers"
      else
        # else its over paragraphs
        looper = Word::IfElseReplacers::IfElseOverParagraphs.new(main_doc, data, options)
        looper.replace_if_else(start_placeholder, end_placeholder, inbetween_placeholders)
      end
    end

    def if_else_are_in_different_container?(start_placeholder, end_placeholder)
      start_placeholder_parent = start_placeholder[:paragraph_object].document
      end_placeholder_parent = end_placeholder[:paragraph_object].document

      start_placeholder_parent != end_placeholder_parent
    end

    def resync_container(container)
      container.parse_paragraphs(container.container_node)
      paragraphs = container.paragraphs
    end

    def there_are_if_else_placeholders?(placeholders)
      placeholders.any?{|p| p[:placeholder_text].include?("if ") }
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
