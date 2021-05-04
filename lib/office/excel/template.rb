require_relative '../constants'
require_relative '../nokogiri_extensions'
require_relative '../excel/placeholder.rb'

# Top-level Excel template rendering.
#
# Replace all {{placeholders}} in all cells of all sheets of the workbook
# template with data, using placeholder syntax like
#
#  {{driver.controller.streams[1].start}}
#  {{driver.controller.streams[1].photo | 133x100}}
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
    class PathNotFound < RuntimeError; end

    module Evaluator
      # Fetch the value from self for the given expr_str, which should be
      # something like "controller.streams[0].start".
      #
      # depends on #dig
      def evaluate(field_path)
        field_path.any? or raise ArgumentError, "Invalid field_path: #{field_path.inspect}"

        # go down field_path, and get the value at each stage.
        # a nil return on the very last step is permitted, otherwise an exception will be raised.
        val, parts_done = field_path.reduce [self, []] do |(obj, path), part|
          break [obj, path] if obj.is_a?(PathNotFound)

          if (rv = obj.dig(part)).nil?
            # Don't double-cuddle {{ }} here otherwise specs break because those cells then look like placeholders again.
            msg = "{#{Office::Placeholder.rejoin field_path}} not found in data"
            msg << " from {#{Office::Placeholder.rejoin path}}" if path.any?
            msg << ?.
            [PathNotFound.new(msg), path << part]
            # raise PathNotFound, "#{part} => nil for path [#{Office::Placeholder.rejoin path}] from [#{Office::Placeholder.rejoin field_path}] on\n #{obj.to_yaml}\n\n from #{self.to_yaml}"
          else
            [rv, path << part]
          end
        end

        if val.is_a?(PathNotFound)
          raise val if field_path != parts_done
        else
          val
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
            placeholder = Office::Placeholder.parse cell.placeholder.to_s
            data.evaluate placeholder.field_path
          rescue
            # TODO maybe use actual xlsx error cells here?
            # TODO maybe have an error callback?
            # TODO test that this doesn't swallow relevant exceptions
            "ERROR: #{$!.message}"
          end

          case val
          # or respond_to? :to_blob
          when Magick::ImageList, Magick::Image
            # add image anchored at this cell
            image_part = sheet.add_image val, cell.location, extent: placeholder.image_extent
            # clear cell value
            # TODO implement delete cell
            cell.value = nil

          when String, Numeric, Date, DateTime, Time, TrueClass, FalseClass, NilClass
            cell.placeholder[] = val.to_s

          when Array
            # TODO insert this during repeat groups work
            tabular = placeholder.options[:tabular]
            cell.value = "Groups not yet implemented, tabular: #{tabular}"

          else
            raise "How to insert #{val.inspect} into sheet?"
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
