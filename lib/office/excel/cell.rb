module Office
  class Cell
    attr_reader :node
    attr_reader :location
    attr_reader :style
    attr_reader :style_id
    attr_reader :data_type
    attr_reader :value_node
    attr_reader :shared_string

    def initialize(c_node, string_table, styles)
      @node = c_node
      @location = c_node["r"]
      @style_id = c_node["s"]
      @data_type = c_node["t"]

      @style = (styles.present? and !@style_id.blank?) ? styles.xf_by_id(@style_id) : nil

      # Originally did this, but was incredibly slow:
      #@value_node = c_node.at_xpath("xmlns:v")
      @value_node = nil
      c_node.elements.each { |e| @value_node = e if e.name == 'v' }

      if is_string? && !@value_node.nil?
        string_id = @value_node.content.to_i
        @shared_string = string_table.get_string_by_id(string_id)
        raise PackageError.new("Excel cell #{@location} refers to invalid shared string #{string_id}") if @shared_string.nil?
        @shared_string.add_cell(self)
      end
    end

    def self.create_node(document, row_number, index, value, string_table)
      cell_node = document.create_element("c")
      cell_node["r"] = "#{column_name(index)}#{row_number}"

      value_node = document.create_element("v")
      cell_node.add_child(value_node)

      unless value.nil? or value.to_s.empty?
        if value.is_a? Numeric
          value_node.content = value
        else
          cell_node["t"] = "s"
          value_node.content = string_table.id_for_text(value.to_s)
        end
      end

      cell_node
    end


    def is_string?
      data_type == "s"
    end

    def self.column_name(index)
      name = ""
      while index >= 0
        name << ('A'.ord + (index % 26)).chr
        index = index/26 - 1
      end
      name.reverse
    end

    def column_num
      letters = /([a-z]+)\d+/i.match(@location)[1].downcase.reverse

      num = letters[0].ord - 'a'.ord
      1.upto(letters.length - 1) { |i| num += (letters[i].ord - 'a'.ord + 1) * (26 ** i) }
      num
    end

    def row_num
      /[a-z]+(\d+)/i.match(@location)[1].to_i - 1
    end

    def value
      return nil if @value_node.nil?
      is_string? ? @shared_string.text : @value_node.content
    end

    def formatted_value
      return nil if @value_node.nil?
      return @shared_string.text if is_string?

      unformatted_value = @value_node.content
      return nil if unformatted_value.nil?

      return unformatted_value if @style.try(:apply_number_format) != '1'

      # ECMA-376 Part 1 section 18.8.30 numFmt (Number Format) - p1767
      case @style.try(:number_format_id)
      when "0"  #    General
        as_decimal(unformatted_value)
      when "1"  #    0
        as_integer(unformatted_value)
      when "2"  #    0.00
        as_decimal(unformatted_value)
      when "3"  #    #,##0
        as_integer(unformatted_value)
      when "4"  #    #,##0.00
        as_decimal(unformatted_value)
      when "9"  #    0%
        as_decimal(unformatted_value)
      when "10" #    0.00%
        as_decimal(unformatted_value)
      when "11" #    0.00E+00
        as_decimal(unformatted_value)
      #when "12" #    # ?/?
      #when "13" #    # ??/??
      when "14" #    mm-dd-yy
        as_date(unformatted_value)
      when "15" #    d-mmm-yy
        as_date(unformatted_value)
      when "16" #    d-mmm
        as_date(unformatted_value)
      when "17" #    mmm-yy
        as_date(unformatted_value)
      when "18" #    h:mm AM/PM
        as_time(unformatted_value)
      when "19" #    h:mm:ss AM/PM
        as_time(unformatted_value)
      when "20" #    h:mm
        as_time(unformatted_value)
      when "21" #    h:mm:ss
        as_time(unformatted_value)
      when "22" #    m/d/yy h:mm
        as_datetime(unformatted_value)
      when "37" #    #,##0 ;(#,##0)
        as_decimal(unformatted_value)
      when "38" #    #,##0 ;[Red](#,##0)
        as_decimal(unformatted_value)
      when "39" #    #,##0.00;(#,##0.00)
        as_decimal(unformatted_value)
      when "40" #    #,##0.00;[Red](#,##0.00)
        as_decimal(unformatted_value)
      when "45" #    mm:ss
        as_time(unformatted_value)
      when "46" #    [h]:mm:ss
        as_time(unformatted_value)
      when "47" #    mmss.0
        as_time(unformatted_value)
      when "48" #    ##0.0E+0
        as_decimal(unformatted_value)
      #when "49" #    @
      else
        unformatted_value
      end
    end

    def as_decimal(value)
      value.to_f
    end

    def as_integer(value)
      value.to_i
    end

    def as_datetime(value)
      DateTime.new(1900, 1, 1, 0, 0, 0) + value.to_f - 2
    end

    def as_date(value)
      DateTime.new(1900, 1, 1, 0, 0, 0) + value.to_i - 2
    end

    def as_time(value)
      as_datetime(value).to_time
    end
  end
end
