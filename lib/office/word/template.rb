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
require 'office/word/placeholder_replacer'

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

    def render(data, options = {})
      paragraph_sets = [main_doc.paragraphs]
      main_doc.headers.each do |header|
        paragraph_sets << header.paragraphs
      end
      main_doc.footers.each do |footer|
        paragraph_sets << footer.paragraphs
      end

      paragraph_sets.each do |paragraphs|
        render_section(paragraphs, data, options)
      end
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
        next_step = {run_index: placeholder[:end_of_placeholder][:run_index], char_index: placeholder[:end_of_placeholder][:char_index] + 1}
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

    def render_section(paragraphs, data, options = {})
      paragraphs.each_with_index do |paragraph, paragraph_index|
        loop_through_placeholders_in_paragraph(paragraph, paragraph_index) do |placeholder|
          replacer = Word::PlaceholderReplacer.new(placeholder, word_document)
          replacement = replacer.replace_in_paragraph(paragraph, data, options)
          next_step = {}
          next_step[:run_index] = replacement[:next_run] if replacement[:next_run]
          next_step[:char_index] = replacement[:next_char] + 1 if replacement[:next_char]
          next_step
        end
      end
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
      run_texts = runs.map(&:text).dup

      next_run_index = 0
      run_texts.each_with_index do |run_text, i|
        next if i < next_run_index
        text = run_text
        next if text.nil?

        next_char_index = 0
        text.each_char.with_index do |char, j|
          next if j < next_char_index
          if char == '{' && next_char(run_texts, i, j)[:char] == '{'
            beginning_of_placeholder = {run_index: i, char_index: j}
            end_of_placeholder = get_end_of_placeholder(run_texts, i, j)
            placeholder_text = get_placeholder_text(run_texts, beginning_of_placeholder, end_of_placeholder)

            placeholder = {placeholder_text: placeholder_text, paragraph_index: paragraph_index, beginning_of_placeholder: beginning_of_placeholder, end_of_placeholder: end_of_placeholder}

            next_step = block_given? ? yield(placeholder) : {}
            if next_step.is_a? Hash
              # This is a bit dodge - even if we increment the run index it will loop through
              # the rest of the chars from the char_index...
              # It doesn't matter because placeholders inside placeholders is not a thing..e but still dodge
              next_run_index = next_step[:run_index] if !next_step[:run_index].nil?
              next_char_index = next_step[:char_index] if !next_step[:char_index].nil?
            end
          end
        end
      end
    end

    def get_end_of_placeholder(run_texts, current_run_index, start_of_placeholder)
      placeholder_text = ""
      start_char = start_of_placeholder
      run_texts[current_run_index..-1].each_with_index do |run_text, i|
        text = run_text
        if !text.nil? && text.length > 0
          text[start_char..-1].each_char.with_index do |char, j|
            the_next_char = next_char(run_texts, current_run_index + i, start_char + j)
            if char == '}' && the_next_char[:char] == '}'
              return {run_index: the_next_char[:run_index], char_index: the_next_char[:char_index]}
            else
              placeholder_text += char
            end
          end
        end
        start_char = 0
      end

      raise InvalidTemplateError.new("Template invalid - end of placeholder }} missing for \"#{placeholder_text}\".")
    end

    def next_char(run_texts, current_run_index, current_char_index)
      current_run_text = run_texts[current_run_index]
      blank = {run_index: nil, char_index: nil, char: nil}
      return blank if current_run_text.nil?

      text = current_run_text || ""
      if text.length - 1 > current_char_index #still chars left at the end
        return {run_index: current_run_index, char_index: current_char_index + 1, char: text[current_char_index + 1]}
      else
        run_texts[current_run_index+1..-1].each_with_index do |run_text, i|
          next if run_text.nil? || run_text.length == 0
          return {run_index: current_run_index+1+i, char_index: 0, char: run_text[0]}
        end
        return blank
      end
    end

    def get_placeholder_text(run_texts, beginning_of_placeholder, end_of_placeholder)
      result = ""
      first_run_index = beginning_of_placeholder[:run_index]
      last_run_index = end_of_placeholder[:run_index]
      if first_run_index == last_run_index
        result = run_texts[first_run_index][beginning_of_placeholder[:char_index]..end_of_placeholder[:char_index]]
      else
        (first_run_index..last_run_index).each do |run_i|
          text = run_texts[run_i]
          next if text.nil? || text.length == 0
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
