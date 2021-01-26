require 'office/word/placeholder_evaluator'

module Word
  class PlaceholderReplacer
    attr_accessor :placeholder, :word_document
    def initialize(placeholder, word_document)
      self.placeholder = placeholder
      self.word_document = word_document
    end

    def replace_in_paragraph(paragraph, data, options = {})
      replacement_result = get_replacement(placeholder, data, options)
      replacement = replacement_result[:replacement]
      render_options = replacement_result[:render_options]

      if options.respond_to?(:deep_dup)
        local_options = options.deep_dup.merge(render_options)
      else
        local_options = options.dup.merge(render_options)
      end

      source_text = placeholder[:placeholder_text]

      # TODO This is identical to code in lib/office/word.rb:82 Office::WordDocument#replace_all
      case
      when replacement.is_a?(String)
        paragraph.replace_all_with_text(source_text, replacement)
      when (replacement.is_a?(Magick::Image) or replacement.is_a?(Magick::ImageList))
        runs = paragraph.replace_all_with_empty_runs(source_text)
        local_options[:document_section] = paragraph.document
        runs.each { |r| r.replace_with_run_fragment(word_document.create_image_run_fragment(replacement, local_options)) }
      else
        runs = paragraph.replace_all_with_empty_runs(source_text)
        runs.each { |r| r.replace_with_body_fragments(word_document.create_body_fragments(replacement, local_options)) }
      end

      {}
    rescue => e
      context_info = "Error in #{placeholder&.dig(:placeholder_text)}"
      raise e.class, [context_info, e.message].join(": ")
    end

    def get_replacement(placeholder, data, options = {})
      evaluator = Word::PlaceholderEvaluator.new(placeholder)
      replacement_settings = evaluator.evaluate(data, options)
      replacement = replacement_settings[:replacement]
      render_options = replacement_settings[:render_options]
      {replacement: replacement, render_options: render_options}
    end
  end
end


__END__

Custom stuff - Might use if the existing replace becomes too slow... but unlikely; its not that much more work

    def replace_in_paragraph(paragraph, data)
      replacement = get_replacement(placeholder, data)
      #source_text = placeholder[:placeholder_text]
      next_step = {}
      case
      when replacement.is_a?(String)
        next_step = replace_text_in_paragraph(paragraph, placeholder, replacement)
      when (replacement.is_a?(Magick::Image) or replacement.is_a?(Magick::ImageList))
        runs = paragraph.replace_all_with_empty_runs(source_text)
        runs.each { |r| r.replace_with_run_fragment(create_image_run_fragment(replacement)) }
      else
        runs = paragraph.replace_all_with_empty_runs(source_text)
        runs.each { |r| r.replace_with_body_fragments(create_body_fragments(replacement, options)) }
      end

      next_step

    end

    def replace_text_in_paragraph(paragraph, placeholder, replacement)
      start_run_index = placeholder[:beginning_of_placeholder][:run_index]
      start_char_index = placeholder[:beginning_of_placeholder][:char_index]

      end_run_index = placeholder[:end_of_placeholder][:run_index]
      end_char_index = placeholder[:end_of_placeholder][:char_index]


      placeholder_length = placeholder[:placeholder_text].to_s.length

      first_run = paragraph.runs[start_run_index]

      if start_run_index == end_run_index
        first_run.text = replace_in_text(first_run.text, start_char_index, placeholder_length, replacement)
        first_run.adjust_for_right_to_left_text
        next_step = {next_run: start_run_index, next_char: start_char_index + placeholder_length}
      else
        length_in_run = first_run.text.length - start_char_index
        first_run.text = replace_in_text(first_run.text, start_char_index, length_in_run, replacement[0,length_in_run])
        first_run.adjust_for_right_to_left_text

        remaining_text = placeholder_length - length_in_run - paragraph.clear_runs((start_run_index + 1), (end_run_index - 1))

        last_run = paragraph.runs[end_run_index]
        last_run.text = replace_in_text(last_run.text, 0, remaining_text, replacement[length_in_run..-1])
        last_run.adjust_for_right_to_left_text

        next_step = {next_run: end_run_index, next_char: remaining_text}
      end

      next_step
    end

    def replace_in_text(original, index, length, replacement)
      return original if length == 0
      result = index == 0 ? "" : original[0, index]
      result += replacement unless replacement.nil?
      result += original[(index + length)..-1] unless index + length == original.length
      result
    end
