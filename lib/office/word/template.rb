=begin
  Templating in word is tricky because runs in paragraphs can start/end randomly.
  So the plan is - Get all the placeholders in the form
  {
    placeholder_text: '{{ i_am_holder }}',
    paragraph_index: 0,
    paragraph_object: <object>,
    begin: {run_index: 3, char_index: 15},
    end: {run_index: 5, char_index: 2}
  }

  Paragraph index mostly just used for when we do the {{ for_each }} stuff

  We can then use the placeholder thingies to render loops and replace the text with the real data.
=end
require 'office/word/placeholder_finder'
require 'office/word/placeholder_replacer'
require 'office/word/for_loop_expander'

class InvalidTemplateError < StandardError
end

module Word
  class Template

    attr_accessor :word_document, :main_doc, :errors



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

    def get_placeholders
      Word::PlaceholderFinder.get_placeholders(main_doc.paragraphs)
    end

    #
    #
    # =>
    # => Rendering
    # =>
    #
    #

    def render(data, options = {})
      paragraph_sets = [main_doc.paragraphs]
      main_doc.headers.each do |header|
        paragraph_sets << header.paragraphs
      end
      main_doc.footers.each do |footer|
        paragraph_sets << footer.paragraphs
      end

      containers = [main_doc, main_doc.headers, main_doc.footers].flatten
      containers.each do |container|

        #TODO - IF_ELSE
        expand_for_loops(container, data, options)
        unless options[:do_not_render] == true
          render_section(container, data, options)
        end
      end
    end

    def expand_for_loops(container, data, options = {})
      paragraphs = container.paragraphs
      expander = Word::ForLoopExpander.new(main_doc, data, options)
      expander.expand_for_loops(container)
    end

    def render_section(container, data, options = {})
      container.paragraphs.each_with_index do |paragraph, paragraph_index|
        Word::PlaceholderFinder.loop_through_placeholders_in_paragraph(paragraph, paragraph_index) do |placeholder|
          replacer = Word::PlaceholderReplacer.new(placeholder, word_document)
          replacement = replacer.replace_in_paragraph(paragraph, data, options)
          next_step = {}
          next_step[:run_index] = replacement[:next_run] if replacement[:next_run]
          next_step[:char_index] = replacement[:next_char] + 1 if replacement[:next_char]
          next_step
        end
      end
    end

    def self.get_value_from_field_identifier(field_identifier, data, options={})
      result = data.with_indifferent_access
      field_recurse = field_identifier.split('.')
      field_recurse.each_with_index do |identifier, i|
        # if result.is_a? Array
        #   result = result.length == 1 ? result.first : {}
        # end

        result = result[identifier]
        if result == nil
          result = ""
          break
        end
      end
      result
    end

  end
end
