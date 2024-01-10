module Word
    class PlaceholderFinderV2
      class << self

        def get_placeholders(paragraphs)
            placeholders = []
            start_time = Time.now
            begin 
              paragraphs.each_with_index do |p, i|
              placeholders += get_placeholders_from_paragraph(p, i)
            end
            end_time = Time.now
            total_time = end_time - start_time
            puts "Total time to get placeholders: #{total_time}"
            byebug
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
            
              matches = text.scan(/(\{\{[^}]*\}\}|\{%[^%]*%\}|\{\s*%[^%]*%\}|\{%[^}]*\}\}|{%\s*if[^%]*%\}|\{%\s*endif\s*%\}|\{%\s*for[^%]*%\}|\{%\s*endfor\s*%\})/) 
                
                text.scan(/(\{\{[^}]*\}\}|\{%[^%]*%\}|\{\s*%[^%]*%\}|\{%[^}]*\}\}|{%\s*if[^%]*%\}|\{%\s*endif\s*%\}|\{%\s*for[^%]*%\}|\{%\s*endfor\s*%\})/)  do |match|
                  placeholder_text = match[0]
                  start_position = Regexp.last_match.begin(0)
                  end_position = Regexp.last_match.end(0) - 1
                  end_char = text[end_position]
                  start_char = text[start_position]
                  last_placeholder = nil
                  last_placeholder = placeholders.last if matches.length > 1 && matches.index(match) > 0 
                
                  placeholders << {
                    placeholder_text: placeholder_text,
                    paragraph_object: {},
                    paragraph_index: paragraph_index,
                    beginning_of_placeholder: {
                      run_index: calculate_run_index(run_texts, start_position),
                      char_index: calculate_beginning_char_index(last_placeholder, run_texts, start_char)
                    },
                    end_of_placeholder: {
                      run_index: calculate_run_index(run_texts, end_position),
                      char_index: calculate_ending_char_index(last_placeholder, run_texts, end_char)
                    }
                  }
              end
            return placeholders
          end

          private

          def calculate_beginning_char_index(last_placeholder, run_texts, start_char)

            current_char_index = nil

            if (last_placeholder.nil?)
              ## TODO: include is nil, so we need to figure out why it's nil or add a nil check. Most likely
              ## will need to do the same for below. 
              found_index = run_texts.index { |str| str.include?(start_char) }
              return found_index || 0
            end


            run_texts.each_with_index do |placeholder_parts, index|
              next unless placeholder_parts.include?(start_char) && index > last_placeholder[:beginning_of_placeholder][:char_index]
              current_char_index = index
              break 
            end

            current_char_index || 0

          end

          def calculate_ending_char_index(last_placeholder, run_texts, end_char)
            current_char_index = 0

            if (last_placeholder.nil?)
              found_index =  run_texts.index{ |str| str.include?(end_char) } 
              
              return found_index || 0
            end

            run_texts.each_with_index do |placeholder_parts, index|
              next unless placeholder_parts.include?(end_char) && index > last_placeholder[:end_of_placeholder][:char_index]
              current_char_index = index
              break 
            end

            current_char_index
      
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