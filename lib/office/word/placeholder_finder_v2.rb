module Word
    class PlaceholderFinderV2
      class << self

        def get_placeholders(paragraphs)
            placeholders = []
            begin 
                paragraphs.each_with_index do |p, i|
                placeholders += get_placeholders_from_paragraph(p, i)
            end
            placeholders
            rescue => e
                  byebug
            end
          end

        def get_placeholders_from_paragraph(paragraph, paragraph_index)
          placeholders = []

          runs = paragraph.runs
          run_texts = runs.map(&:text).dup

          return [] if run_texts.empty? || run_texts.nil?
              text = run_texts.join('')
            
              text.scan(/(\{\{[^}]*\}\}|\{%[^%]*%\}|\{\s*%[^%]*%\}|\{%[^}]*\}\}|{%\s*if[^%]*%\}|\{%\s*endif\s*%\}|\{%\s*for[^%]*%\}|\{%\s*endfor\s*%\})/) do |match|
                  placeholder_text = match[0]
                  
                  start_position = Regexp.last_match.begin(0)
                  end_position = Regexp.last_match.end(0) - 1
                  end_char = text[end_position]
                  start_char = text[start_position]
      
                  start_position_run_index = calculate_run_index(run_texts, start_position)
                  start_position_char_index = run_texts[start_position_run_index].index(start_char)
                  end_position_run_index = calculate_run_index(run_texts, end_position)
                  end_position_char_index = run_texts[end_position_run_index].index(end_char)

                  placeholders << {
                    placeholder_text: placeholder_text,
                    paragraph_object: paragraph,
                    paragraph_index: paragraph_index,
                    beginning_of_placeholder: {
                      run_index: start_position_run_index,
                      char_index: start_position_char_index,
                    },
                    end_of_placeholder: {
                      run_index: end_position_run_index,
                      char_index: end_position_char_index,
                    }
                  }
              end
            return placeholders
          end

          private

          def calculate_run_index(run_texts, position)
            current_position = 0
            run_texts.each_with_index do |text, run_index|
              text = text || ""
              return run_index if current_position + text.length > position
              current_position += text.length
            end
            run_texts.size - 1
          end
      end
    end
end