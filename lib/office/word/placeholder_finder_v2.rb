module Word
    class PlaceholderFinderV2
      class << self

        def get_placeholders(paragraphs)
            placeholders = []
        
                paragraphs.each_with_index do |p, i|
                placeholders += get_placeholders_from_paragraph(p, i)
            end
            placeholders
          
          end
        
        def get_placeholders_from_paragraph(paragraph, paragraph_index)
          placeholders = []
          previous_run_hash = {}

          runs = paragraph.runs
          run_texts = runs.map(&:text).dup

          return [] if run_texts.empty? || run_texts.nil?
              text = run_texts.join('')
              check_brace_balance(text)
            
              text.scan(/(\{\{[^}]*\}\}|\{%[^%]*%\}|\{\s*%[^%]*%\}|\{%[^}]*\}\}|{%\s*if[^%]*%\}|\{%\s*endif\s*%\}|\{%\s*for[^%]*%\}|\{%\s*endfor\s*%\})/) do |match|
                  placeholder_text = match[0]
                  
                  start_position = Regexp.last_match.begin(0)
                  start_char = text[start_position]

                  end_position = Regexp.last_match.end(0) - 1
                  end_char = text[end_position]

                  beginning_of_placeholder = get_placeholder_positions(run_texts, start_position, start_char, previous_run_hash, "start")
                  end_of_placeholder = get_placeholder_positions(run_texts, end_position, end_char, previous_run_hash, "end")
                 
                  placeholders << {
                    placeholder_text: placeholder_text,
                    paragraph_object: paragraph,
                    paragraph_index: paragraph_index,
                    beginning_of_placeholder: {
                      run_index: beginning_of_placeholder[:run_index],
                      char_index: beginning_of_placeholder[:char_index],
                    },
                    end_of_placeholder: {
                      run_index: end_of_placeholder[:run_index],
                      char_index: end_of_placeholder[:char_index],
                    }
                  }
              end
            return placeholders
          end

          private

          def generate_identifier(start_or_end, position_run_index)
            start_or_end == "start" ? "S-#{position_run_index}" : "E-#{position_run_index}"
          end
        
          def get_placeholder_positions(run_texts, position, passed_char, previous_run_hash, start_or_end)
            position_run_index = calculate_run_index(run_texts, position)
            position_char_index = run_texts[position_run_index].index(passed_char)
            identifier = generate_identifier(start_or_end, position_run_index)
            hash_key = start_or_end == "start" ? "used_start_indexes" : "used_end_indexes"
          
            if previous_run_hash.key?(identifier)
              ignore_indexes = previous_run_hash[identifier][hash_key.to_sym]
              run_text = run_texts[position_run_index]
              run_text.length.times do |index|
                char = run_text[index]
                next if ignore_indexes.include?(index)
          
                next_char = run_texts[position_run_index][position_char_index + 1]&.chr
          
                if passed_char == next_char && (char == "}" || char == "%")
                  position_char_index = index + 1
                  break
                end
          
                if char == passed_char
                  position_char_index = index
                  break
                end
              end
          
              previous_run_hash[identifier][hash_key.to_sym] << position_char_index
            else
              if start_or_end == "end"
                next_char = run_texts[position_run_index][position_char_index + 1]&.chr
                position_char_index += 1 if next_char == passed_char
              end
              
              previous_run_hash[identifier] = { hash_key.to_sym => [position_char_index] }
            end
          
            { char_index: position_char_index, run_index: position_run_index, previous_run_hash: previous_run_hash }
          end
          

          def check_brace_balance(text)
            unbalanced_occurrences = text.scan(/{{[^{}]*[^{}]*$/)
            if unbalanced_occurrences.any?
              raise InvalidTemplateError.new("Template invalid - end of placeholder }} missing for \"#{unbalanced_occurrences.first}\".")
            end
          end

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