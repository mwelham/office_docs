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
            tabular = placeholder.options[:tabular]
            tabular_data =
            case val.first
            when Array
              # assume array of arrays
              val
            when Hash
              table val
            end

            sheet.accept!(cell.location, tabular_data)

          else
            raise "How to insert #{val.inspect} into sheet?"
          end
        end
      end

      workbook
    end

    # Convert a tabular array (ie [field_names, *records]) to an
    # array of {field_name => value} hashes, one for each record.
    #
    # 30-Apr-2021 Used in specs only.
    module_function def tabular_hashify(tabular_array)
      # split tabular data
      field_names, *records = tabular_array

      # convert each row to a hash
      records.map {|ary| field_names.zip(ary).to_h }
    end

    # This is proof-of-concept rather than working code. Leaving it here
    # because it's easier to understand than distribute.
    #
    # path is a set of keys, leaf is index for now (augment with row later?)
    # paths is an accumulator - the map from each possible path to its index
    # last index will be somewhere in paths, so track it separately
    module_function def path_indices node, path = [], paths = {}, last_index = 0
      case node
      when Hash
        # for each name/value in the hash:
        # see if paths[path + [name]] has an index
        # if not add it and increment last_index
        node.reduce [paths, last_index] do |(paths, last_index), (name, value)|
          extended_path = path + [name]
          path_indices value, extended_path, paths, last_index
        end
      when Array
        node.reduce [paths, last_index] do |(paths, last_index), child_node|
          path_indices child_node, path, paths, last_index
        end
      else
        # singular value, so increment index if necessary
        incremented_index = paths[path] ||= last_index + 1
        [paths, incremented_index]
      end
    end

    module_function def group_hash node
      empty = Hash.new{|h,k| h[k] = {}}
      node.each_with_object empty do |(name,value), groups|
        value_type =
        case value
        # when ->_{Array === value && value.size == 1}; Hash
        when Array; Array
        when Hash; Hash
        else Object
        end

        empty[value_type][name] = value
      end
    end

    module_function def distribute node, row_so_far = [], path = [], paths = {}, last_index = nil, &blk
      # for each name/value in the hash:
      # see if paths[path + [name]] has an index
      # if not add it and increment last_index
      case node
      when Hash
        value_types = group_hash node

        # collect Object/singular values to construct prefix
        paths, last_index, prefix = value_types[Object].reduce([paths, last_index, row_so_far]) do |(paths, last_index, row), (name, value)|
          # no blk passed here because augmented prefix will come back as return value
          distribute value, row, (path + [name]), paths, last_index do |*| end
        end

        # process hash values, which may or may not contain array children

        # process array values.
        # TODO ? Effectively an array of size 1 is the same as a hash value
        hashes, arrays = value_types[Hash], value_types[Array]
        case [hashes.size, arrays.size]
        when [0, 0]
          # This is at a node with only leafs, so yield as a row.
          blk.call prefix
          # also return as a row. Because. Maybe its useful.
          [paths, last_index, prefix]
        else
          # put hashes (ie things containing prefixes) before arrays
          kids = hashes.merge arrays
          # accumulate separate rows rather than accumulating in one
          kids.reduce([paths, last_index, []]) do |(paths, last_index, rows), (name, value)|
            npaths, index, row = distribute value, prefix, (path + [name]), paths, last_index, &blk
            [npaths, index, (rows + [row])]
          end
        end

      when Array
        # necessarily returns an array of rows
        node.reduce([paths, last_index, []]) do |(paths, last_index, rows), child_node|
          paths, index, row = distribute child_node, row_so_far, path, paths, last_index, &blk
          [paths, index, (rows + [row])]
        end

      else
        # singular value, so increment index if necessary
        unless index = paths[path]
          index = paths[path] = last_index ? last_index + 1 : 0
        end
        # put value in the correct place in the array
        # TODO .dup here is quite inefficient.
        (new_row = row_so_far.dup)[index] = node
        [paths, index, new_row]
      end
    end

    module_function def table node
      rows = []
      # distribute the values from the tree into a rectangular format
      paths, index, nested_rows = distribute node do |row| rows << row end
      # pad each row to maximum length in case we want to transpose
      rows.each{|r| r[index] ||= nil}
      # make headers from dotted paths
      headers = paths.sort_by{|_,index| index}.map{|name,_| name.join ?.}
      # put headers with
      rows.unshift headers
    end
  end
end
