require 'office/word/for_loop_expanders/base'

module Word
  module ForLoopExpanders
    class LoopInParagraph < Word::ForLoopExpanders::Base

      # Break runs on start + end
      # loop runs inbetween

      def expand_loop(start_placeholder, end_placeholder, inbetween_placeholders)
        paragraph = start_placeholder[:paragraph_object]

        start_run = paragraph.replace_first_with_empty_runs(start_placeholder[:placeholder_text]).last
        end_run = paragraph.replace_first_with_empty_runs(end_placeholder[:placeholder_text]).first

        from_run = paragraph.runs.index(start_run) + 1
        to_run = paragraph.runs.index(end_run) - 1

        inbetween_runs = paragraph.runs[from_run..to_run]
        for_loop_placeholder_info = parse_for_loop_placeholder(start_placeholder[:placeholder_text])
        #This is the 0 run of our loop.
        # Get a set of runs for the duplicate runs
        duplicate_runs = generate_new_run_set(paragraph, inbetween_runs)


        field_data = for_loop_placeholder_info[:data].presence || []
        if field_data.blank?
          inbetween_runs.each do |run|
            paragraph.remove_run(run)
          end
        else
          replace_variable_in_placeholders(0, for_loop_placeholder_info, inbetween_placeholders, paragraph, inbetween_runs)
          (field_data[1..-1] || []).each_with_index do |data_set, i|
            new_run_set = generate_new_run_set(paragraph, duplicate_runs)
            replace_variable_in_placeholders(i+1, for_loop_placeholder_info, inbetween_placeholders, paragraph, new_run_set)

            new_run_set.each do |run|
              paragraph.add_new_run_object_before_run(run, end_run)
            end
          end
        end

      end

      def generate_new_run_set(paragraph, runs)
        new_run_set = []
        runs.each do |r|
          new_run_set << Office::Run.new(r.node.clone, paragraph)
        end
        new_run_set
      end
    end
  end
end
