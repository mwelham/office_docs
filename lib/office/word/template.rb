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
require 'office/word/if_else_replacer'

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
      containers = [main_doc, main_doc.headers, main_doc.footers].flatten
      containers.each do |container|
        expand_for_loops(container, data, options)
        replace_if_else(container, data, options)
        unless options[:do_not_render] == true
          render_section(container, data, options)
        end
      end

      fixBrokenPRThings(containers)
    end

    def expand_for_loops(container, data, options = {})
      paragraphs = container.paragraphs
      expander = Word::ForLoopExpander.new(main_doc, data, options)
      expander.expand_for_loops(container)
    end

    def replace_if_else(container, data, options = {})
      paragraphs = container.paragraphs
      expander = Word::IfElseReplacer.new(main_doc, data, options)
      expander.replace_all_if_else(container)
    end

    def render_section(container, data, options = {})
      container.paragraphs.each_with_index do |paragraph, paragraph_index|
        Word::PlaceholderFinder.get_placeholders_from_paragraph(paragraph, paragraph_index) do |placeholder|
          replacer = Word::PlaceholderReplacer.new(placeholder, word_document)
          replacement = replacer.replace_in_paragraph(paragraph, data, options)
          next_step = {}
          next_step[:run_index] = replacement[:next_run] if replacement[:next_run]
          next_step[:char_index] = replacement[:next_char] + 1 if replacement[:next_char]
          next_step
        end
      end
    end

    # This is some weirdness - but word gets angry if there are duplicate docPr objects.
    # These come from graphic objects like lines.
    # When word repairs the files it just increments them... so we are just doing that...
    def fixBrokenPRThings(containers)
      trouble_nodes = []
      used_ids = []
      containers.each do |container|
        trouble_nodes += container.xml_node.xpath('//wp:docPr[@id!=""]')
        used_ids += container.xml_node.xpath('//*[@id][@id!=""]').map{|n| n["id"].to_i}
      end
      current_id = (used_ids.max || 0) + 1
      trouble_nodes.group_by{|n| n["id"]}.each do |id, nodes|
        next if nodes.count <= 1
        nodes[1..-1].each do |n|
          n["id"] = current_id
          n["name"] = n["name"].gsub(id.to_s, current_id.to_s)
          current_id += 1
        end
      end
    end

    def self.get_value_from_field_identifier(field_identifier, data, options={})
      result = data.with_indifferent_access
      field_recurse = field_identifier.split('.')
      field_recurse.each_with_index do |identifier, i|
        # This part is if are trying to get a value from an array - because result is still the last result
        # e.g. ArrayAnswer.Billy
        # If there is only 1 thing we allow them to access it without array syntax
        if result.is_a? Array
          result = result.length == 1 ? result.first : {}
        end

        array_info_from_identifier = parse_array_info_from_identifier(identifier)
        identifier = array_info_from_identifier[:identifier_without_array_info]
        result = result[identifier]

        if array_info_from_identifier[:index] != nil
          index = array_info_from_identifier[:index]
          result = result[index] if result.is_a? Array
        end

        if result.nil?
          result = ""
          break
        end
      end
      result
    end

    def self.parse_array_info_from_identifier(identifier)
      result = identifier.match(/.+\[(.+)\]/)

      new_identifier = identifier
      index = nil

      if result
        index = result[1].to_i
        new_identifier = identifier.gsub(/\[(.+)\]/,'')
      end

      {index: index, identifier_without_array_info: new_identifier}
    end

    def self.remove_node(node)
      parent = node.parent
      if node.name == 'p' && parent.name == 'tc' && parent.children.select{|c| c.name == 'p'}.count == 1
        node.content = ""
      else
        node.remove
      end
    end

  end
end
