module Word
  class ForLoopExpander
    attr_accessor :main_doc, :paragraphs, :data, :options, :placeholders
    def initialize(main_doc, paragraphs, data, options = {})
      self.main_doc = main_doc
      self.paragraphs = paragraphs
      self.data = data
      self.options = options
    end

    def expand_for_loops
      # Get placeholders in paragraphs
      self.placeholders = Word::PlaceholderFinder.get_placeholders(paragraphs)
      i = 0
      while i < placeholders.length
        puts i
        start_placeholder = placeholders[i]
        if start_placeholder[:placeholder_text].include?("foreach")
          end_index = get_end_index(i)
          expand_loop(i, end_index)

          i = end_index + 1
        else
          i += 1
        end
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
        loop_inside_paragraph(start_placeholder, end_placeholder, inbetween_placeholders)
        # Break runs on start + end
        # loop runs inbetween
      elsif false
        # else if start is in a table cell and end is in a different table cell
        # loop whole row
      elsif false
        # else if start is in a table cell but end is not in a table cell at all
        # raise error
      else
        # else
        # break paragraphs on start + end
        # loop paragraphs inbetween
      end
    end

    def loop_inside_paragraph(start_placeholder, end_placeholder, inbetween_placeholders)
      paragraph = start_placeholder[:paragraph_object]

      start_run = paragraph.replace_first_with_empty_runs(start_placeholder[:placeholder_text]).last
      end_run = paragraph.replace_first_with_empty_runs(end_placeholder[:placeholder_text]).first

      from_run = paragraph.runs.index(start_run) + 1
      to_run = paragraph.runs.index(end_run) - 1

      inbetween_runs = paragraph.runs[from_run..to_run]
      #This is the 0 run of our loop.
      # Get a set of text for the duplicate runs
      #run_texts = inbetween_runs.map(&:text)
      for_loop_placeholder_info = parse_for_loop_placeholder(start_placeholder[:placeholder_text])

      field_data = for_loop_placeholder_info[:data].presence || []
      field_data[1..-1].each_with_index do |data_set, i|
        new_run_set = generate_new_run_set(paragraph, inbetween_runs)
        replace_variable_in_placeholders(i+1, for_loop_placeholder_info, inbetween_placeholders, paragraph, new_run_set)

        new_run_set.each do |run|
          paragraph.add_new_run_object_before_run(run, end_run)
        end
      end
      #Replace placeholders with extrapolated placeholders
      replace_variable_in_placeholders(0, for_loop_placeholder_info, inbetween_placeholders, paragraph, inbetween_runs)
    end

    def generate_new_run_set(paragraph, runs)
      new_run_set = []
      runs.each do |r|
        new_run_set << Office::Run.new(r.node.clone, paragraph)
      end
      new_run_set
    end

    def replace_variable_in_placeholders(index, for_loop_placeholder_info, placeholders, paragraph, inbetween_runs)
      placeholders.each do |p|
        placeholder_variable_matcher = /#{for_loop_placeholder_info[:variable]}\./
        placeholder = p[:placeholder_text]
        if placeholder.match(placeholder_variable_matcher)
          new_placeholder = placeholder.gsub(placeholder_variable_matcher,"lolpies[#{index}]")
          paragraph.replace_all_with_text(placeholder, new_placeholder, inbetween_runs)
        end
      end
    end

    def parse_for_loop_placeholder(placeholder)
      result = placeholder.gsub('{%','').gsub('%}','').match(/foreach (\w+) in (.+)/)
      variable = result[1].strip
      data_pointer = result[2].strip
      raise "Invalid syntax for foreach placeholder #{placeholder}" if variable.blank? || data_pointer.blank?
      field_data = Word::Template.get_value_from_field_identifier(data_pointer, data)
      {variable: variable, data_pointer: data_pointer, data: field_data}
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
