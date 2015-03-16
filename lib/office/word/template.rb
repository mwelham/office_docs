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

    attr_accessor :errors

    class InvalidTemplateError < StandardError
    end

    attr_accessor :word_document, :main_doc
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

      placeholders = get_placeholders(paragraphs)

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

      runs = paragraph.runs

      next_run_index = 0
      runs.each_with_index do |run, i|
        next if i < next_run_index
        text = run.text

        next_char_index = 0
        text.each_char.with_index do |char, j|
          next if j < next_char_index
          if char == '{' and next_char(runs, i, j)[:char] == '{'
            #We have found the start of a placeholder!
            beginning_of_placeholder = {run_index: i, char_index: j}
            end_of_placeholder = get_end_of_placeholder(runs, i, j)
            placeholder_text = get_placeholder_text(runs, beginning_of_placeholder, end_of_placeholder)

            placeholders << {placeholder: placeholder_text, paragraph_index: paragraph_index, beginning_of_placeholder: beginning_of_placeholder, end_of_placeholder: end_of_placeholder}

            #Skip ahead to the end of this placeholder
            next_run_index = end_of_placeholder[:run_index]
            next_char_index = end_of_placeholder[:char_index] + 1
          end
        end
      end

      placeholders
    end

    def get_end_of_placeholder(runs, current_run_index, start_of_placeholder)
      start_char = start_of_placeholder
      runs[current_run_index..-1].each_with_index do |run, i|
        text = run.text
        text[start_char..-1].each_char.with_index do |char, j|
          the_next_char = next_char(runs, current_run_index + i, j)
          if char == '}' && the_next_char[:char] == '}'
            return {run_index: current_run_index + i, char_index: the_next_char[:char_index]}
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

    #
    #
    # =>
    # => Rendering
    # =>
    #
    #

    def render_section(paragraphs, data)
      placeholders = get_placeholders(paragraphs)

      paragraphs.each do |p|
        template_paragraph(p, data)
      end
    end









  end
end
