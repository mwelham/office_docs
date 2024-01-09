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
            placeholders
            rescue => e
                byebug
            end
          end

        def get_placeholders_from_paragraph(paragraph, paragraph_index)
            placeholders = []

            runs = paragraph.runs
            run_texts = runs.map(&:text).dup

            #TODO: Should this be nil, or do we need to return an object of some sort?
            return [] if run_texts.empty? || run_texts.nil?

            run_texts.each do |text|
                text = text || ""
                text.scan(/(\{\{[^}]*\}\}|\{%[^%]*%\}|\{\{[^%]*%\}|\{%[^}]*\}\})/) do |match|
                    placeholder_text = match[0]
                    start_position = Regexp.last_match.begin(0)
                    end_position = Regexp.last_match.end(0) - 1
            
                    return placeholders << {
                        placeholder_text: placeholder_text,
                        paragraph_object: paragraph,
                        paragraph_index: paragraph_index,
                        beginning_of_placeholder: {
                         run_index: calculate_run_index(run_texts, start_position),
                         char_index: calculate_char_index(text, start_position)
                        },
                        end_of_placeholder: {
                         run_index: calculate_run_index(run_texts, end_position),
                         char_index: calculate_char_index(text, end_position)
                        }
                    }
                end
                return placeholders
            end
            return placeholders
        end

          private

          def calculate_run_index(run_texts, position)
            current_position = 0
            run_texts.each_with_index do |text, run_index|
              return run_index if current_position + text.length > position
              current_position += text.length
            end
            # Default to the last run if the position exceeds the total length of runs
            run_texts.size - 1
          end
          
    
          def calculate_char_index(paragraph, position)
            # Logic to calculate char_index based on position
            paragraph.lines.last.nil? ? position : position % paragraph.lines.last.length
          end
      end
    end
end