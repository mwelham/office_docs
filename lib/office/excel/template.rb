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
              table_of_hash val
            end

            sheet.accept!(cell.location, tabular_data)

          else
            raise "How to insert #{val.inspect} into sheet?"
          end
        end
      end

      workbook
    end

    # convert an array of hashes to a table of rows
    module_function def table_of_hash hash
      headers = [make_hash_counter]
      Array tablify hash, headers
    end

    # Convert a tabular array (ie [field_names, *records]) to an
    # array of {field_name => value} hashes, one for each record.
    module_function def tabular_hashify(tabular_array)
      # split tabular data
      field_names, *records = tabular_array

      # convert each row to a hash
      records.map {|ary| field_names.zip(ary).to_h }
    end

    # headers track possibly varying headers across different child blocks
    module_function def tablify singular_value_hashes, headers, &blk
      return enum_for __method__, singular_value_hashes, headers unless block_given?

      # not all hashes contain the same keys, so we need to keep track, always
      # put the same fields in the same positions, and append fields that show
      # up for the first time.
      values = singular_value_hashes.map do |singular_value_hash|
        field_values = []
        singular_value_hash.each do |field,value|
          field_index = headers.last[field]
          field_values[field_index] = value
        end
        field_values
      end

      # yield rows
      values.each &blk
    end

    module_function def make_hash_counter
      Hash.new{|h,k| h[k] = h.size}
    end

    Row = Struct.new :row
    Header = Struct.new :header
    Block = Struct.new :headers, :rows

    # convert headers to largest
    # headers = ary.select{|(t,*r)| t == :header}.map{|ry| ry.last.flat_map{|headers| headers&.keys || [nil]}}.max_by(&:length)
    # rows = ary.select{|t,_| t == :row}.map{|(_t,r)|r.flatten}
    # Convert a nested has of values and array to an
    # array of arrays with repeated data
    # TODO keep track of headers (which may not be unique between levels)
    # headers should probably be an array of header => pos, one per level
    module_function def table_of(node, prefix: [], headers: [make_hash_counter], &blk)
      # return enum_for __method__, node, prefix: prefix, headers: headers unless block_given?

      case node
      when Array
        rows = node.map do |child|
          table_of child, prefix: prefix, headers: headers, &blk
        end

        # to signal child end-of-block
        blk&.call :header, headers

        if rows.all?{|r| Row === r}
          Block.new headers, rows
        else
          rows
        end

      when Hash
        # separate values (vals) from arrays (kids)
        kids, vals = node.partition{|_k,v| Array === v || Hash === v}

        # wrap in an Array to unpack it to values on next recursive call
        nprefix = tablify([vals.to_h], headers).to_a

        if kids.empty?
          # no further nesting so yield this whole row
          appended_row = prefix + nprefix
          blk&.call :row, appended_row
          Row.new appended_row
          # {headers: headers, rows: [appended_row]}
        else
          # more nesting, so yield this prefix with each nested row
          # yield child array with locally-appended prefix
          headers = [*headers, {key: 0}, make_hash_counter]
          # to accumulate/reduce headers do
          # headers << nil << make_hash_counter
          rows = kids.map do |(key, node)|
            # yield data rows individually
            table_of node, prefix: [*prefix, *nprefix, [key]], headers: headers, &blk
          end
          # {headers: headers, children: rows}
          rows.flatten
        end
      else
        node
      end
    end

    # convert a node in a data hierarchy to a set of header/data rows.
    # Doesn't quite work because the prefix goes awry somewhere
    module_function def to_row_blocks node
      ary = table_of node
      grouped_by_header = ary.group_by{|block| block.headers}
      rows = []
      grouped_by_header.each do |header,blocks|
        # binding.pry unless $dont
        rows << header.flat_map(&:keys)
        blocks.each do |block|
          block.rows.each do |r|
            rows << r.row.flatten
          end
        end
      end
      rows
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

    module_function def distribute node, row_so_far = [], path = [], paths = {}, last_index = nil
      # for each name/value in the hash:
      # see if paths[path + [name]] has an index
      # if not add it and increment last_index
      case node
      when Hash
        # TODO to_h is not necessary, just change param destructuring
        kids, vals = node.partition{|_k,v| Array === v || Hash === v}.map(&:to_h)

        # collect singular values
        # NOTE dup is necessary to copy the prefix to each result row, otherwise
        # it gets overwritten by the singular value assignment, below.
        paths, last_index, prefix = vals.reduce([paths, last_index, row_so_far.dup]) do |(paths, last_index, row), (name, value)|
          distribute value, row, (path + [name]), paths, last_index
        end

        # process array values, ie child nodes, if they exist
        paths, last_index, row =
        if kids.any?
          # accumulate separate rows rather than accumulating in one
          paths, last_index, row = kids.reduce([paths, last_index, []]) do |(paths, last_index, rows), (name, value)|
            npaths, index, row = distribute value, prefix, (path + [name]), paths, last_index
            [npaths, index, (rows + row)]
          end
          [paths, last_index, row.flatten]
        else
          # This is an actual row, not a recursive walk artifact. So wrap it in
          # a non-array to protect against flattening.
          [paths, last_index, Row.new(prefix)]
        end

        [paths, last_index, row]
      when Array
        # necessarily returns an array of rows
        node.reduce([paths, last_index, []]) do |(paths, last_index, rows), child_node|
          paths, index, row = distribute child_node, row_so_far, path, paths, last_index
          [paths, index, (rows + [row])]
        end
      else
        # singular value, so increment index if necessary
        unless index = paths[path]
          index = paths[path] = last_index ? last_index + 1 : 0
        end
        row_so_far[index] = node
        [paths, index, row_so_far]
      end
    end
  end
end
