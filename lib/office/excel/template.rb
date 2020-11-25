require_relative '../excel.rb'
require_relative '../constants'
require_relative '../nokogiri_extensions'

# Top-level Excel template rendering.
#
# Replace all {{placeholders}} in all cells of all sheets of the workbook
# template with data, using placeholder syntax like
#
#  {{driver.controller.streams[1].start}}
#
# data must be a HoHA (Hash of Hashes and Arrays), also known as 'a json' :-p
#
# Example use case
#
#  workbook = Office::ExcelWorkbook.new('workbook-template.xlsx')
#  data = JSON.parse(some_json_string, symbolize_names: true)
#  Excel::Template.render!(workbook, data)
#  workbook.save("workbook-#{data[:name]}.xlsx")
#
module Excel
  # 'module' because there's no state, and therefore no justification for
  # creating an instance to hold said non-existent state.
  module Template
    # %r because / breaks my editor's syntax highlighting with the double } :-s
    PLACEHOLDER_RX = %r|{{.*?}}|

    module Evaluator
      # Fetch the value from self for the given expr_str, which should be
      # something like "controller.streams[0].start".
      #
      # depends on #dig
      #
      # TODO this will probably end up being too naive, because we might want
      # more information about which prefix evaluated to nil.
      def evaluate(expr_str)
        # split on . and [] so "streams[0].start" becomes [:streams, 0, :start]
        parts = expr_str.split(/[\[.\]]+/).map! do |part|
          # convert to integer, then symbol
          Integer part rescue part.to_sym
        end

        dig *parts
      rescue ArgumentError => ex
        # check that we're getting the exception from dig and not somewhere else
        if ex.message == 'wrong number of arguments (given 0, expected 1+)'
          raise "Invalid expression: #{expr_str.inspect}"
        else
          raise
        end
      end
    end

    # This is the non-destructive render, so it will return a new ExcelWorkbook
    # leaving workbook untouched.
    #
    # see render! for other ways data can conform.
    module_function def render(workbook, data)
      workbook.clone.tap do |target|
        render!(target, data)
      end
    end

    # Renders values from data into placeholders in workbook.
    # NOTE modifies workbook.
    #
    # Returns modified workbook as a convenience
    #
    # data can be a Hash, or it can also an object that understands #evaluate
    # (see module Evaluator), or failing that an object that understands #dig.
    module_function def render!(workbook, data)
      # don't modify original 'data' object if it already has evaluate method
      unless data.respond_to? :evaluate
        data = data.dup.extend Evaluator
      end

      # evaluate placeholders on all sheets
      workbook.sheets.each do |sheet|
        sheet.each_placeholder do |cell|
          val =
          begin
            data.evaluate(cell.placeholder)
          rescue
            # TODO maybe use actual xlsx error cells here?
            # TODO maybe have an error callback?
            "ERROR for cell.placeholder: #{$!.message}"
          end

          cell.value =
          if cell.value.length > cell.placeholder.length + 4
            # cell contains text surrounding {{placeholder}}
            # TODO maybe this is Cell responsibility?
            # TODO maybe use the word placeholder code here?
            # NOTE we can't rely on Excel formatting of value here, so just convert to a string
            cell.value.gsub(PLACEHOLDER_RX, val.to_s)
          else
            val
          end
        end
      end

      workbook
    end

    # Convert a tabular array (ie [field_names, *records]) to an
    # array of {field_name => value} hashes, one for each record.
    module_function def tabular_hashify(tabular_array)
      # split tabular data
      field_names, *records = tabular_array

      # convert each row to a hash
      records.map {|ary| field_names.zip(ary).to_h }
    end
  end
end
