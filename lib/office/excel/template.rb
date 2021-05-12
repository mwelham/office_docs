require_relative '../constants'
require_relative '../nokogiri_extensions'
require_relative '../excel/placeholder.rb'

# Top-level Excel template rendering.
#
# Replace all {{placeholders}} in all cells of all sheets of the workbook
# template with data, using placeholder syntax like
#
# For singular values
#  {{driver.controller.streams[1].start}}
#
# For images
#  {{driver.controller.streams[1].photo | 133x100}}
#
# For tabular data
#  {{driver.controller.streams | tabular,horizontal,headers,insert }}
#  {{driver.controller.streams | tabular,vertical,headers }}
#  {{driver.controller.streams | tabular,headers }}
#  {{driver.controller.streams}}
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

          when Array, Hash
            render_tabular sheet, cell, placeholder, val

          else
            raise "How to insert #{val.inspect} : #{val.class} into sheet?"
          end
        end
      end

      workbook
    end

    # write values into the area specified by cell, placeholder and placeholder options
    # source_data is either an Array of Hashes, or a Hash
    # returns the Office::Range of the data written
    module_function def render_tabular sheet, cell, placeholder, source_data
      tabular_data, max_index = table source_data, **placeholder.options.slice(:tabular, :headers)

      # transpose rows depending on options
      tabular_data, width, height = maybe_transpose tabular_data, max_index, **placeholder.options.slice(:vertical, :horizontal)

      if placeholder.options[:insert]
        insert_range = cell.location * [width-1, height-1]
        sheet.insert_rows insert_range
      end

      # write values to cells
      sheet.accept!(cell.location, tabular_data)
    end

    # Convert a tabular array (ie [field_names, *records]) to an
    # array of {field_name => value} hashes, one for each record.
    #
    # 30-Apr-2021 Used in specs only. But budget is tight so not fixing it.
    module_function def tabular_hashify(tabular_array)
      # split tabular data
      field_names, *records = tabular_array

      # convert each row to a hash
      records.map {|ary| field_names.zip(ary).to_h }
    end

    # This is proof-of-concept rather than working code. Leaving it here
    # because it's easier to understand than distribute.
    #
    # node is a Hash of Hashes and Arrays collection aka 'a json'. ie it's a tree.
    #
    # path is a set of keys which constitutes a path in node. The leaf of that path
    # will be an index in the output row.
    #
    # paths is the accumulator - the map from each possible path to its index in
    # an output array.
    #
    # last_index will be buried somewhere in paths, so track it separately
    #
    # So this method will traverse node and produce a map of
    #
    #  path => index_in_output_row
    #
    # so that values from sub-trees can be grouped together without their
    # indices overlapping with other subtrees.
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

    # node is a hash of names to values.
    # do a group_by of the kind of value (ie one of Array, Hash, Object)
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

    # For each name/value in the hash:
    #
    # - calculate index of value in row using paths[path + [name]], if not exist
    #   add it and increment last_index
    #
    # - yield a row (index is calculated above) whenever there is a collection
    #   of singular values only, ie every key has a non-array and non-hash value.
    #
    # Uses a slightly unusual combination of return value (to collect the
    # singular values) and yielding a full row (to avoid nested arrays
    # containing rows). As a kinda pleasant bonus, top-level return value will
    # contain the nested rows.
    #
    # See comments for path_indices for a description of parameters.
    module_function def distribute node, row_so_far = [], path = [], paths = {}, last_index = nil, &blk
      case node
      when Hash
        value_types = group_hash node

        # prefix is constructed from row_so_far with collected Object/singular values
        # It's recursive to keep the index calculation going
        paths, last_index, prefix = value_types[Object].reduce([paths, last_index, row_so_far]) do |(paths, last_index, row), (name, value)|
          # no blk passed here because augmented prefix will come back as return value
          distribute value, row, (path + [name]), paths, last_index do |*| end
        end

        # Process hash values, which may or may not contain array children.
        # Then process array values.
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
          # traverse child nodes
          kids.reduce([paths, last_index, []]) do |(paths, last_index, rows), (name, value)|
            # NOTE path extended with current key
            npaths, index, row = distribute value, prefix, (path + [name]), paths, last_index, &blk
            [npaths, index, (rows + [row])]
          end
        end

      when Array
        # necessarily returns an array of rows
        # expects child_node to be a hash
        node.reduce([paths, last_index, []]) do |(paths, last_index, rows), child_node|
          raise "cannot handle non-hash in #{__method__}" unless child_node.is_a? Hash

          # NOTE uses parent's non-extended path
          npaths, index, row = distribute child_node, row_so_far, path, paths, last_index, &blk
          [npaths, index, (rows + [row])]
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

    # rows is an array of arrays
    # max_index is the longest array in rows
    # returns rows, width, height
    #
    # vertical is the default
    # vertical means field names across the top
    # horizontal means field names on the left
    module_function def maybe_transpose rows, max_index, horizontal: false, vertical: false
      # optional orientation
      # would be clearer with ruby-2.7 pattern matching
      orientation =
      case [vertical, horizontal]
      when [true, true];   :vertical # tie-breaker case
      when [true, false];  :vertical   # specified
      when [false, true];  :horizontal # specified
      when [false, false]; :vertical # default case
      else
        raise "unknown orientations: #{{vertical: vertical, horizontal: horizontal}.inspect}"
      end

      case orientation
      when :vertical
        [rows, max_index, rows.size]

      when :horizontal
        # pad each row to maximum length so we can transpose
        [rows.each{|r| r[max_index] ||= nil}.transpose, rows.size, max_index]

      else
        raise "unknown orientation: #{orientation}"
      end
    end

    # returns [rows, max_index] where max_index is the length of the longest
    # array in rows
    #
    # tabular: is just to absorb the value coming in
    module_function def table node, tabular: nil, headers: false
      rows = []

      # distribute the values from the tree into a rectangular format
      paths, max_index, nested_rows = distribute node do |row| rows << row end

      # optional headers
      if headers
        # make headers from dotted paths
        headers = paths.sort_by{|_,index| index}.map{|name,_| name.join ?.}
        rows.unshift headers
      end

      [rows, max_index]
    end
  end
end
