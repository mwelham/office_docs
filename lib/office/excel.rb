require 'office/package'
require 'office/constants'
require 'office/errors'
require 'office/logger'

require_relative 'excel/cell'
require_relative 'excel/excel_workbook'
require_relative 'excel/sheet'
require_relative 'excel/template'

module Office
  class Row
    attr_reader :node
    attr_reader :number
    attr_reader :spans
    attr_reader :string_table, :styles

    def initialize(row_node, string_table, styles)
      @node = row_node

      @number = Integer(row_node[:r]) - 1
      @spans = row_node[:spans]
      @string_table = string_table
      @styles = styles
    end

    def cells
      @cells ||= node.
        xpath("#{Package.xpath_ns_prefix(node)}:c").
        map{ |c| Cell.new(c, string_table, styles) }
    end

    def self.create_node(document, number, data, string_table, styles:)
      row_node = document.create_element("row")
      row_node["r"] = number.to_s unless number.nil?

      unless data.nil? or data.length == 0
        row_node["spans"] = "1:#{data.length}"
        0.upto(data.length - 1) do |i|
          c_node = Cell.create_node(document, number, i, data[i], string_table, styles: styles)
          row_node.add_child(c_node)
        end
      end

      row_node
    end

    def cells_padded_to_column_number
      ary = []
      cells.each do |c|
        ary.push(nil) until ary.length > c.location.coli
        ary[c.location.coli] = c
      end
      ary
    end

    def to_ary
      # TODO why might c be nil here?, and surely you'd want to keep it nil rather than ''?
      cells_padded_to_column_number.map { |c| c&.value || '' }
    end
  end

  class SharedStringTable
    attr_reader :node

    def initialize(part)
      ns_prefix = Package.xpath_ns_prefix(part.xml.root)
      @node = part.xml.at_xpath("/#{ns_prefix}:sst")

      # TODO Keep these up-to-date
      @count_attr = @node.attribute("count")
      @unique_count_attr = @node.attribute("uniqueCount")

      @strings_by_id = {}
      @strings_by_text = {}
      node.xpath("#{ns_prefix}:si").each { |si| parse_si_node(si) }
    end

    def parse_si_node(si)
      string = SharedString.new(si, @strings_by_id.length)
      @strings_by_id[string.id] = string
      @strings_by_text[string.text] = string
      string.id
    end

    def get_string_by_id(id)
      @strings_by_id[id]
    end

    def id_for_text(text)
      return @strings_by_text[text].id if @strings_by_text.has_key? text

      si = node.document.create_element("si")
      t = node.document.create_element("t")
      t.content = text
      si.add_child(t)
      @node.add_child(si)

      parse_si_node(si)
    end

    def debug_dump
      rows = @strings_by_id.values.collect do |s|
        cells = s.cells.collect { |c| c.location }
        ["#{s.id}", "#{s.text}", "#{cells.join(', ')}"]
      end

      footer = ","
      footer << "  count = #{@count_attr.value}" unless @count_attr.nil?
      footer << "  unique count = #{@unique_count_attr.value}" unless @unique_count_attr.nil?
      Logger.debug_dump_table("Excel Workbook Shared Strings", ["ID", "Text", "Cells"], rows, footer)
    end
  end

  class SharedString
    attr_reader :node
    attr_reader :text_node
    attr_reader :id

    # Nothing in the gem uses this. Which indicates the whole class may be kinda pointless.
    # attr_reader :cells

    def initialize(si_node, id)
      @node = si_node
      @id = id
      # // is slower but we need to handle text runs
      @text_node = si_node.xpath(".//#{Package.xpath_ns_prefix(si_node)}:t")
      @cells = []
    end

    def text
      text_node&.text
    end

    def add_cell(cell)
      @cells << cell
    end
  end

  class StyleSheet
    # From section 18.8.30 of the specs - these are the predefined format numFmt ids
    KNOWN_FORMAT_IDS = [0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37, 38, 39, 40, 45, 46, 47, 48, 49]

    attr_reader :node, :xfs

    def initialize(part)
      @node = part.xml.nxpath('/*:styleSheet').first
      @xfs = node.nxpath('*:cellXfs/*:xf').map do |xf_node|
        CellXF.new(xf_node)
      end
    end

    def xf_by_index(index)
      @xfs[Integer index]
    end

    def create_xf num_fmt_id
    end

    def index_of_xf(num_fmt_id)
      ix = @xfs.index{|cell_xf| cell_xf.number_format_id == num_fmt_id}
      if ix.nil? && KNOWN_FORMAT_IDS.include?(num_fmt_id)
        # This xml is from LibreOffice Calc, so hopefully it works for Google Docs, Excel etc.
        cell_xfs_node = @node.nxpath('*:cellXfs').first
        # TODO what should fontID here be?
        cell_xfs_node << %|<xf numFmtId="#{num_fmt_id}" applyNumberFormat="1" fontId="0" fillId="0" borderId="0" xfId="0"/>|
        # append new CellXF node
        @xfs << CellXF.new(cell_xfs_node.element_children.last)
        cell_xfs_node[:count] = @xfs.size
        # index of the last one added
        ix = @xfs.size - 1
      end
      ix
    end
  end

  # Cell style
  class CellXF
    attr_reader :node
    attr_reader :number_format_id
    attr_reader :apply_number_format

    def apply_number_format?
      @apply_number_format == 1
    end

    def initialize(xf_node)
      @node = xf_node
      @xf_id = Integer xf_node['xfId']
      @number_format_id = Integer xf_node['numFmtId']
      @apply_number_format = xf_node['applyNumberFormat'].to_i
    end
  end
end
