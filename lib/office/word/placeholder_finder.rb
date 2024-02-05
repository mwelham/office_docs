module Word
  class PlaceholderFinder
    class << self

      START_OF_PLACEHOLDER = 'START'  
      START_IDENTIFIER_PREFIX = "S"

      END_IDENTIFIER_PREFIX = "E"
      END_OF_PLACEHOLDER = 'END'
      
        # This regex is used to find placeholders in the following format:
        # {{...}} - variable placeholder
        # {% if ... %} - liquid syntax placeholder
        # {% endif %} - liquid syntax placeholder
        # {% for ... %} - liquid syntax placeholder
        # {% endfor %} - liquid syntax placeholder

      PLACEHOLDER_REGEX = /(\{\{[^}]*\}\}|\{%[^%]*%\}|\{%[^}]*\}\}|{%\s*(if|endif|for|endfor)[^%]*%\})/
      UNBALANCED_PLACEHOLDER_BRACES_REGEX = /{{[^{}]*[^{}]*$/
      
      def get_placeholders(paragraphs)
        start_time = Time.now
        placeholders = []
        paragraphs.each_with_index do |p, i|
          placeholders += get_placeholders_from_paragraph(p, i)
        end
        end_time = Time.now
        puts "Time to get placeholders: #{end_time - start_time}"
        placeholders
      end
      
      def get_placeholders_from_paragraph(paragraph, paragraph_index)
        placeholders = []
        previous_run_hash = {}
        runs = paragraph.runs
        run_texts = runs.map(&:text)
        
        return [] if run_texts.empty? || run_texts.nil?
        text = run_texts.join('')
        check_brace_balance(text)
          
        text.scan(PLACEHOLDER_REGEX) do |match|
          placeholder_text = match[0]
          
          start_position = Regexp.last_match.begin(0)
          start_char = text[start_position]

          end_position = Regexp.last_match.end(0) - 1
          end_char = text[end_position]

          # This is used to get the char_index & run_index of the placeholder in the run_texts array
          beginning_of_placeholder = get_placeholder_positions(run_texts, start_position, start_char, previous_run_hash, START_OF_PLACEHOLDER)
          end_of_placeholder = get_placeholder_positions(run_texts, end_position, end_char, previous_run_hash, END_OF_PLACEHOLDER)
          
          placeholders << {
            placeholder_text: placeholder_text,
            paragraph_object: paragraph, #TODO: Remove this to save memory, and create a method to get the paragraph object from the index
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

        # The identifier is used to identify the start/end placeholder in the previous_run_hash
        # due to nested placeholders, we store the used indexes in the previous_run_hash 
        # and use them to skip over them when searching for the next placeholder start/end index
        # E - end placeholder, S - start placeholder
        def generate_identifier(start_or_end, position_run_index)
          start_or_end == START_OF_PLACEHOLDER ? "#{START_IDENTIFIER_PREFIX}-#{position_run_index}" : "#{END_IDENTIFIER_PREFIX}-#{position_run_index}"
        end

        # Ex: run_texts = ["Nane:" , "{{", "field.name", "}}", "Age:", "{{", "field.age", "}}"]
        # run_index is the start/end index of the placeholder in the run_texts array
        # char_index is the index of the start/end of the string in the run_texts array
        # (think of run_index and char_index as coordinates)
        #
        # Placeholder: {{ field.name }}
        # Beginning of placeholder: {  - run_index: 1, char_index: 0  
        # End of placeholder: } - run_index: 3, char_index: 1 
        #
        # Placeholder: {{ field.age }}
        # Beginning of placeholder: {  - run_index: 5, char_index: 0
        # End of placeholder: } - run_index: 7, char_index: 1
      
        def get_placeholder_positions(run_texts, position, passed_char, previous_run_hash, start_or_end)
          position_run_index = calculate_run_index(run_texts, position)
          position_char_index = run_texts[position_run_index].index(passed_char)
          identifier = generate_identifier(start_or_end, position_run_index)
          
          # If the previous_run_hash has the identifier, it means that there are multiple placeholders in the same run, i.e same index of
          # the runs array. We use the previous_run_hash to skip over the indexes that have already been used.
          # we do this to ensure we get the correct start/end index of the placeholder 
          if previous_run_hash.key?(identifier)
            ignore_indexes = previous_run_hash[identifier] || []
            run_text = run_texts[position_run_index]
            run_text.length.times do |index|
              # Skip if we already visited this index
              next if ignore_indexes.include?(index)
              char = run_text[index]

              # If the char is a { or %, we check if the next char is the same as the passed_char
              # If it is and it's the end of the placeholder we found the correct index and use that. 
              # ex: ending placeholder - {{...}} - passed_char is } and next_char is } - we want to use passed_char + 1 to get the correct ending index
              next_char = run_texts[position_run_index][index + 1]&.chr 
              
              if passed_char == next_char && (char == "}" || char == "%") && start_or_end == END_OF_PLACEHOLDER && !ignore_indexes.include?(index + 1)
                position_char_index = index + 1
                previous_run_hash[identifier] << index
                previous_run_hash[identifier] << position_char_index
                break
              end
              
              if char == passed_char
                position_char_index = index
                previous_run_hash[identifier] << index
                # if we are at the start of the placeholder, we want to store the index of the next char { or % t since it's used. 
                if start_or_end == START_OF_PLACEHOLDER
                  previous_run_hash[identifier] << (index + 1)
                end

                break
              end
              # don't check this index again
              previous_run_hash[identifier] << index
            end
            # We add the index to the previous_run_hash so we can skip over it if we find another placeholder in the same run
           
          else
            if start_or_end == END_OF_PLACEHOLDER
              # If we are at the end of the placeholder, we want to grab the index of the last char of the placeholder
              # ex: {{...}} - we want to grab the index of the last } in the placeholder
              next_char = run_texts[position_run_index][position_char_index + 1]&.chr
              position_char_index += 1 if next_char == passed_char
            end
            
            previous_run_hash[identifier] = [position_char_index] 
            previous_run_hash[identifier] << (position_char_index + 1)
          end
        
          { char_index: position_char_index, run_index: position_run_index, previous_run_hash: previous_run_hash }
        end
        
        # This method is used to check if placeholder braces are properly closed in the template
        # in the following format: {{...}} is valid but {{...} and {{ ... are not valid. 
        def check_brace_balance(text)
          unbalanced_occurrences = text.scan(UNBALANCED_PLACEHOLDER_BRACES_REGEX)
          if unbalanced_occurrences.any?
            raise InvalidTemplateError.new("Template invalid - end of placeholder }} missing for \"#{unbalanced_occurrences.first}\".")
          end
        end

        # This method is used to calculate the run index of the placeholder
        # Run index is the index of the placeholder in run_texts array
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