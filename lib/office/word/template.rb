=begin
  Templating in word is tricky because runs in paragraphs can start/end randomly.
  So the plan is - Get all the placeholders in the form
  {
    placeholder: '{{ i_am_holder }}',
    paragraph_index: 0,
    begin: {run_index: 3, char_index: 15},
    end: {run_index: 5, char_index: 2}
  }

  Paragraph index mostly just used for when we do the {{ for_each }} stuff

  We can then use the placeholder thingies to render loops and replace the text with the real data.
=end
module Word
  class Template

    attr_accessor :word_document, :main_doc, :errors

    class InvalidTemplateError < StandardError
    end


    def initialize(word_document)
      self.word_document = word_document
      self.main_doc = word_document.main_doc
    end

    def template_valid?
      begin
        get_placeholders
      rescue InvalidTemplateError => e
        self.errors = e.message
        return false
      end
      return true
    end

    def render(data)
      paragraphs = main_doc.paragraphs #Create various sections using the #for_each later
      render_section(paragraphs, data)
    end

    #
    #
    # =>
    # => Getting placeholders
    # =>
    #
    #

    def get_placeholders(paragraphs = main_doc.paragraphs)
      placeholders = []
      paragraphs.each_with_index do |p, i|
        placeholders += get_placeholders_from_paragraph(p, i)
      end
      placeholders
    end

    def get_placeholders_from_paragraph(paragraph, paragraph_index)
      placeholders = []
      loop_through_placeholders_in_paragraph(paragraph, paragraph_index) do |placeholder|
        placeholders << placeholder
        {run_index: placeholder[:end_of_placeholder][:run_index], char_index: placeholder[:end_of_placeholder][:char_index] + 1}
      end
      placeholders
    end


    #
    #
    # =>
    # => Rendering
    # =>
    #
    #

    def render_section(paragraphs, data)
      paragraphs.each_with_index do |paragraph, paragraph_index|
        loop_through_placeholders_in_paragraph(paragraph, paragraph_index) do |placeholder|
          replacement = replace_in_paragraph(paragraph, placeholder, data)
          {run_index: replacement[:end_run], char_index: replacement[:end_char] + 1}
        end
      end
    end

    def replace_in_paragraph(paragraph, placeholder, data)
      start_run_index = placeholder[:beginning_of_placeholder][:run_index]
      start_char_index = placeholder[:beginning_of_placeholder][:char_index]

      end_run_index = placeholder[:end_of_placeholder][:run_index]
      end_char_index = placeholder[:end_of_placeholder][:char_index]

      replacement = get_replacement(placeholder, data)
      placeholder_length = placeholder[:placeholder].to_s.length

      first_run = paragraph.runs[start_run_index]
      index_in_run = start_char_index

      if start_run_index == end_run_index
        first_run.text = replace_in_text(first_run.text, index_in_run, placeholder_length, replacement)
        first_run.adjust_for_right_to_left_text
        result = {end_run: start_run_index, end_char: index_in_run + placeholder_length}
      else
        length_in_run = first_run.text.length - index_in_run
        first_run.text = replace_in_text(first_run.text, index_in_run, length_in_run, replacement[0,length_in_run])
        first_run.adjust_for_right_to_left_text

        remaining_text = placeholder_length - length_in_run - paragraph.clear_runs((start_run_index + 1), (end_run_index - 1))

        last_run = paragraph.runs[end_run_index]
        last_run.text = replace_in_text(last_run.text, 0, remaining_text, replacement[length_in_run..-1])
        last_run.adjust_for_right_to_left_text

        result = {end_run: end_run_index, end_char: remaining_text}
      end

      result

    end

    def get_replacement(placeholder, data)
      placeholder_text = placeholder[:placeholder_text]
      #TODO: Evaluate placeholder and options - work out replacement
      replacement = "urka durka mohammed jihaad"
    end

    def replace_in_text(original, index, length, replacement)
      return original if length == 0
      result = index == 0 ? "" : original[0, index]
      result += replacement unless replacement.nil?
      result += original[(index + length)..-1] unless index + length == original.length
      result
    end


    #
    #
    # =>
    # => Magical placeholder looping stuff
    # =>
    #
    #


    def loop_through_placeholders_in_paragraph(paragraph, paragraph_index)
      runs = paragraph.runs

      next_run_index = 0
      runs.each_with_index do |run, i|
        next if i < next_run_index
        text = run.text

        next_char_index = 0
        text.each_char.with_index do |char, j|
          next if j < next_char_index
          if char == '{' and next_char(runs, i, j)[:char] == '{'
            beginning_of_placeholder = {run_index: i, char_index: j}
            end_of_placeholder = get_end_of_placeholder(runs, i, j)
            placeholder_text = get_placeholder_text(runs, beginning_of_placeholder, end_of_placeholder)

            placeholder = {placeholder: placeholder_text, paragraph_index: paragraph_index, beginning_of_placeholder: beginning_of_placeholder, end_of_placeholder: end_of_placeholder}
            next_step = block_given? ? yield(placeholder) : {}
            if next_step.is_a? Hash
              next_run_index = next_step[:run_index] if !next_step[:run_index].nil?
              next_char_index = next_step[:char_index] if !next_step[:char_index].nil?
            end
          end
        end
      end
    end

    def get_end_of_placeholder(runs, current_run_index, start_of_placeholder)
      start_char = start_of_placeholder
      runs[current_run_index..-1].each_with_index do |run, i|
        text = run.text
        text[start_char..-1].each_char.with_index do |char, j|
          the_next_char = next_char(runs, current_run_index + i, start_char + j)
          if char == '}' && the_next_char[:char] == '}'
            return {run_index: the_next_char[:run_index], char_index: the_next_char[:char_index]}
          end
        end
        start_char = 0
      end

      raise InvalidTemplateError.new("Template invalid - end of placeholder }} missing.")
    end

    def next_char(runs, current_run_index, current_char_index)
      current_run = runs[current_run_index]
      text = current_run.text
      if text.length - 1 > current_char_index #still chars left at the end
        return {run_index: current_run_index, char_index: current_char_index + 1, char: text[current_char_index + 1]}
      else
        runs[current_run_index+1..-1].each_with_index do |run, i|
          next if run.text.length == 0
          return {run_index: current_run_index+1+i, char_index: 0, char: run.text[0]}
        end
        return {run_index: nil, char_index: nil, char: nil}
      end
    end

    def get_placeholder_text(runs, beginning_of_placeholder, end_of_placeholder)
      result = ""
      first_run_index = beginning_of_placeholder[:run_index]
      last_run_index = end_of_placeholder[:run_index]
      if first_run_index == last_run_index
        result = runs[first_run_index].text[beginning_of_placeholder[:char_index]..end_of_placeholder[:char_index]]
      else
        (first_run_index..last_run_index).each do |run_i|
          text = runs[run_i].text
          if run_i == first_run_index
            result += text[beginning_of_placeholder[:char_index]..-1]
          elsif run_i == last_run_index
            result += text[0..end_of_placeholder[:char_index]]
          else
            result += text
          end
        end
      end
      result
    end

  end
end
