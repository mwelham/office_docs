#encoding: UTF-8

require 'office/package'
require 'office/constants'
require 'office/errors'
require 'office/logger'

module Office
  class WordDocument < Package
    attr_accessor :main_doc
    
    def initialize(filename)
      super(filename)

      main_doc_part = get_relationship_targets(WORD_MAIN_DOCUMENT_TYPE).first
      raise PackageError.new("Word document package '#{@filename}' has no main document part") if main_doc_part.nil?
      @main_doc = MainDocument.new(self, main_doc_part)
    end

    def self.blank_document
      doc = WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'blank.docx'))
      doc.filename = nil
      doc
    end

    def self.from_data(data)
      file = Tempfile.new('OfficeWordDocument')
      file.binmode
      file.write(data)
      file.close
      begin
        doc = WordDocument.new(file.path)
        doc.filename = nil
        return doc
      ensure
        file.delete
      end
    end

    def add_heading(text)
      p = @main_doc.add_paragraph
      p.add_style("Heading1")
      p.add_text_run(text)
      p
    end

    def add_sub_heading(text)
      p = @main_doc.add_paragraph
      p.add_style("Heading2")
      p.add_text_run(text)
      p
    end

    def add_paragraph(text)
      p = @main_doc.add_paragraph
      p.add_text_run(text)
      p
    end

    def add_image(image) # image must be an Magick::Image or ImageList
      p = @main_doc.add_paragraph
      p.add_run_with_fragment(create_image_run_fragment(image))
      p
    end

    def add_table(hash, options = {}) # keys of hash are column headings, each value an array of column data
      @main_doc.add_table(create_table_fragment(hash, options))
    end

    def plain_text
      @main_doc.plain_text
    end

    # The type of 'replacement' determines what replaces the source text:
    #   Image  - an image (Magick::Image or Magick::ImageList)
    #   Hash   - a table, keys being column headings, and each value an array of column data
    #   Array  - a sequence of these replacement types all of which will be inserted
    #   String - simple text replacement
    def replace_all(source_text, replacement, options = {})
      case
      # For simple cases we just replace runs to try and keep formatting/layout of source
      when replacement.is_a?(String)
        @main_doc.replace_all_with_text(source_text, replacement)
      when (replacement.is_a?(Magick::Image) or replacement.is_a?(Magick::ImageList))
        runs = @main_doc.replace_all_with_empty_runs(source_text)
        runs.each { |r| r.replace_with_run_fragment(create_image_run_fragment(replacement)) }
      else
        runs = @main_doc.replace_all_with_empty_runs(source_text)
        runs.each { |r| r.replace_with_body_fragments(create_body_fragments(replacement, options)) }
      end
    end

    def create_body_fragments(item, options = {})
      case
      when (item.is_a?(Magick::Image) or item.is_a?(Magick::ImageList))
        [ "<w:p>#{create_image_run_fragment(item)}</w:p>" ]
      when item.is_a?(Hash)
        [ create_table_fragment(item, options) ]
      when item.is_a?(Array)
        create_multiple_fragments(item, options)
      else
        [ create_paragraph_fragment(item.nil? ? "" : item.to_s) ]
      end
    end

    def create_image_run_fragment(image)
      prefix = ["", @main_doc.part.path_components, "media", "image"].flatten.join('/')
      identifier = unused_part_identifier(prefix)
      extension = "#{image.format}".downcase

      part = add_part("#{prefix}#{identifier}.#{extension}", StringIO.new(image.to_blob), image.mime_type)
      relationship_id = @main_doc.part.add_relationship(part, IMAGE_RELATIONSHIP_TYPE)

      Run.create_image_fragment(identifier, image.columns, image.rows, relationship_id)
    end

    def create_table_fragment(hash, options = {})
      c_count = hash.size
      return "" if c_count == 0

      c_index = 0
      fragment = "<w:tbl>#{create_table_properties_fragment(c_count, options)}<w:tr>"
      hash.keys.each do |header|
        column_properties = create_column_properties_fragment(c_index, c_count, options)
        c_index += 1
        encoded_header = Nokogiri::XML::Document.new.encode_special_chars(header.to_s)
        fragment << "<w:tc>#{column_properties}<w:p><w:r><w:t>#{encoded_header}</w:t></w:r></w:p></w:tc>"
      end
      fragment << "</w:tr>"

      r_count = hash.values.inject(0) { |max, value| [max, value.is_a?(Array) ? value.length : (value.nil? ? 0 : 1)].max }
      0.upto(r_count - 1).each do |i|
        fragment << "<w:tr>"
        hash.values.each { |v| fragment << create_table_cell_fragment(v, i) }
        fragment << "</w:tr>"
      end

      fragment << "</w:tbl>"
      fragment
    end

    MAX_TABLE_WIDTH_IN_TWIPS = 10592

    # Valid options for customizing the table are:
    #   options[:table_style]     : the name of the style to use
    #   options[:use_full_width]  : set to true to explicitly declare column sizes to try and fill the page width
    def create_table_properties_fragment(column_count, options)
      # If the 'LightGrid' style is not present in the original Word doc (it is with our blank) then the style is ignored.
      style = (options.nil? or options[:table_style].nil? or options[:table_style].empty?) ? "LightGrid" : options[:table_style]

      properties = '<w:tblPr>'
      properties << "<w:tblStyle w:val=\"#{Nokogiri::XML::Document.new.encode_special_chars(style.to_s)}\"/>"
      properties << '<w:tblW w:w="0" w:type="auto"/>'
      properties << '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
      properties << '</w:tblPr>'

      if autosize_columns?(options)
        properties << '<w:tblGrid/>'
      else
        properties << '<w:tblGrid>'
        (column_count - 1).times { |i| properties << "<w:gridCol w:w=\"#{column_width(i, column_count)}\"/>" }
        properties << "<w:gridCol w:w=\"#{column_width(column_count - 1, column_count)}\"/>"
        properties << '</w:tblGrid>'        
      end

      properties
    end

    def create_column_properties_fragment(zero_based_index, count, options)
      return "" if autosize_columns?(options)

      properties = '<w:tcPr>'
      properties << "<w:tcW w:w=\"#{column_width(zero_based_index, count)}\" w:type=\"dxa\"/>"
      properties << '</w:tcPr>'
      properties
    end

    def autosize_columns?(options)
      (options.nil? or options[:use_full_width] != true)
    end

    def column_width(zero_based_index, count)
      if zero_based_index < (count - 1)
        (MAX_TABLE_WIDTH_IN_TWIPS / count).to_i
      else
        MAX_TABLE_WIDTH_IN_TWIPS - (MAX_TABLE_WIDTH_IN_TWIPS / count).to_i * (count - 1)
      end
    end

    def create_table_cell_fragment(values, index)
      item = case
      when (!values.is_a?(Array))
        index != 0 || values.nil? ? "" : values
      when index < values.length
        values[index]
      else
        ""
      end

      xml = create_body_fragments(item).join
      # Word vaildation rules seem to require a w:p immediately before a /w:tc
      xml << "<w:p/>" unless xml.end_with?("<w:p/>") or xml.end_with?("</w:p>")
      "<w:tc>#{xml}</w:tc>"
    end

    def create_multiple_fragments(array, options = {})
      array.map { |item| create_body_fragments(item, options) }.flatten
    end

    def create_paragraph_fragment(text)
      "<w:p><w:r><w:t>#{Nokogiri::XML::Document.new.encode_special_chars(text)}</w:t></w:r></w:p>"
    end

    def debug_dump
      super
      @main_doc.debug_dump
      #Logger.debug_dump_xml("Word Main Document", @main_doc.part.xml)
    end
  end

  class ParagraphContainer
    attr_accessor :container_node
    attr_accessor :paragraphs

    def parse_paragraphs(node)
      @container_node = node
      @paragraphs = []
      node.xpath(".//w:p").each { |p| @paragraphs << Paragraph.new(p, self) }
    end

    def add_paragraph
      p_node = @container_node.add_child(@container_node.document.create_element("p"))
      @paragraphs << Paragraph.new(p_node, self)
      @paragraphs.last
    end

    def paragraph_inserted_after(existing, additional)
      p_index = @paragraphs.index(existing)
      raise ArgumentError.new("Cannot find paragraph after which new one was inserted") if p_index.nil?

      @paragraphs.insert(p_index + 1, additional)
    end

    def add_table(xml_fragment)
      table_node = @container_node.add_child(xml_fragment)
      table_node.xpath(".//w:p").each { |p| @paragraphs << Paragraph.new(p, self) }
    end

    def plain_text
      text = ""
      @paragraphs.each do |p| 
        p.runs.each { |r| text << r.text unless r.text.nil? }
        text << "\n"
      end
      text
    end

    def replace_all_with_text(source_text, replacement_text)
      @paragraphs.each { |p| p.replace_all_with_text(source_text, replacement_text) }
    end

    def replace_all_with_empty_runs(source_text)
      @paragraphs.collect { |p| p.replace_all_with_empty_runs(source_text) }.flatten
    end

    def debug_stats
      stats = { :p_count => 0, :r_count => 0, :t_chars => 0 }
      @paragraphs.each do |p|
        stats[:p_count] += 1
        p.runs.each do |r|
          stats[:r_count] += 1
          stats[:t_chars] += r.text_length
        end
      end
      stats
    end

    def debug_dump_stats(part_name)
      stats = debug_stats
      Logger.debug_dump "#{part_name} Stats"
      Logger.debug_dump "  paragraphs  : #{stats[:p_count]}"
      Logger.debug_dump "  runs        : #{stats[:r_count]}"
      Logger.debug_dump "  text length : #{stats[:t_chars]}"
      Logger.debug_dump ""
    end

    def debug_dump_plain_text(part_name)
      Logger.debug_dump "#{part_name} Plain Text"
      Logger.debug_dump ">>>"
      Logger.debug_dump plain_text
      Logger.debug_dump "<<<"
      Logger.debug_dump ""
    end
  end

  class MainDocument < ParagraphContainer
    attr_accessor :part
    attr_accessor :body_node
    attr_accessor :headers
    attr_accessor :footers
    
    def initialize(word_doc, part)
      @parent = word_doc
      @part = part
      parse_xml
      parse_headers
      parse_footers
    end
    
    def parse_xml
      xml_doc = @part.xml
      @body_node = xml_doc.at_xpath("/w:document/w:body")
      raise PackageError.new("Word document '#{@filename}' is missing main document body") if @body_node.nil?
      parse_paragraphs(@body_node)
    end

    def parse_headers
      @headers = @part.get_relationship_targets(DOCX_HEADER_TYPE).map { |part| Header.new(part, self) }
    end

    def parse_footers
      @footers = @part.get_relationship_targets(DOCX_FOOTER_TYPE).map { |part| Footer.new(part, self) }
    end

    def replace_all_with_text(source_text, replacement_text)
      super
      @headers.each { |h| h.replace_all_with_text(source_text, replacement_text) }
      @footers.each { |f| f.replace_all_with_text(source_text, replacement_text) }
    end

    def replace_all_with_empty_runs(source_text)
      runs = super
      @headers.each { |h| runs += h.replace_all_with_empty_runs(source_text) }
      @footers.each { |f| runs += f.replace_all_with_empty_runs(source_text) }
      runs
    end

    def debug_dump
      debug_dump_stats("Main Document")
      debug_dump_plain_text("Main Document")

      if @headers.empty?
        Logger.debug_dump "(no headers present for document)"
        Logger.debug_dump ""
      else
        @headers.each_index do |i|
          @headers[i].debug_dump_stats("Header #{i + 1}")
          @headers[i].debug_dump_plain_text("Header #{i + 1}")
        end
      end
        
      if @footers.empty?
        Logger.debug_dump "(no footers present for document)"
        Logger.debug_dump ""
      else
        @footers.each_index do |i|
          @footers[i].debug_dump_stats("Footer #{i + 1}")
          @footers[i].debug_dump_plain_text("Footer #{i + 1}")
        end
      end
    end
  end

  class Header < ParagraphContainer
    attr_accessor :part
    attr_accessor :main_doc
    attr_accessor :header_node

    def initialize(header_part, parent_doc)
      @part = header_part
      @main_doc = parent_doc
 
      @header_node = part.xml.at_xpath("/w:hdr")
      raise PackageError.new("Word document '#{@filename}' is missing hdr root in header XML") if @header_node.nil?
      parse_paragraphs(@header_node)
     end
  end

  class Footer < ParagraphContainer
    attr_accessor :part
    attr_accessor :main_doc
    attr_accessor :footer_node

    def initialize(footer_part, parent_doc)
      @part = footer_part
      @main_doc = parent_doc

      @footer_node = part.xml.at_xpath("/w:ftr")
      raise PackageError.new("Word document '#{@filename}' is missing ftr root in footer XML") if @footer_node.nil?
      parse_paragraphs(@footer_node)
    end
  end

  class Paragraph
    attr_accessor :node
    attr_accessor :runs
    attr_accessor :document
    
    def initialize(p_node, parent)
      @node = p_node
      @document = parent
      @runs = []
      p_node.xpath("w:r").each { |r| @runs << Run.new(r, self) }
    end

    # TODO Wrap styles up in a class
    def add_style(style)
      pPr_node = @node.add_child(@node.document.create_element("pPr"))
      pStyle_node = pPr_node.add_child(@node.document.create_element("pStyle"))
      pStyle_node["w:val"] = style
      # TODO return style object
    end

    def add_text_run(text)
      r_node = @node.add_child(@node.document.create_element("r"))
      populate_r_node(r_node, text)

      r = Run.new(r_node, self)
      @runs << r
      r
    end

    def populate_r_node(r_node, text)
      t_node = r_node.add_child(@node.document.create_element("t"))
      t_node["xml:space"] = "preserve"
      t_node.content = text
    end

    def add_run_with_fragment(fragment)
      r = Run.new(@node.add_child(fragment), self)
      @runs << r
      r
    end

    def replace_all_with_text(source_text, replacement_text)
      return if source_text.nil? or source_text.empty?
      replacement_text = "" if replacement_text.nil?

      text = @runs.inject("") { |t, run| t + (run.text || "") }
      until (i = text.index(source_text, i.nil? ? 0 : i)).nil?
        replace_in_runs(i, source_text.length, replacement_text)
        text = replace_in_text(text, i, source_text.length, replacement_text)
        i += replacement_text.length
      end
    end
    
    def replace_all_with_empty_runs(source_text)
      return [] if source_text.nil? or source_text.empty?

      empty_runs = []
      text = @runs.inject("") { |t, run| t + (run.text || "") }
      until (i = text.index(source_text, i.nil? ? 0 : i)).nil?
        empty_runs << replace_with_empty_run(i, source_text.length)
        text = replace_in_text(text, i, source_text.length, "")
      end
      empty_runs
    end

    def replace_with_empty_run(index, length)
      replaced = replace_in_runs(index, length, "")
      first_run = replaced[0]
      index_in_run = replaced[1]

      r_node = @node.document.create_element("r")
      run = Run.new(r_node, self)
      case
      when index_in_run == 0
        # Insert empty run before first_run
        first_run.node.add_previous_sibling(r_node)
        @runs.insert(@runs.index(first_run), run)
      when index_in_run == first_run.text.length
        # Insert empty run after first_run
        first_run.node.add_next_sibling(r_node)
        @runs.insert(@runs.index(first_run) + 1, run)
      else
        # Split first_run and insert inside
        preceding_r_node = @node.add_child(@node.document.create_element("r"))
        populate_r_node(preceding_r_node, first_run.text[0..index_in_run - 1])
        first_run.text = first_run.text[index_in_run..-1]

        first_run.node.add_previous_sibling(preceding_r_node)
        @runs.insert(@runs.index(first_run), Run.new(preceding_r_node, self))

        first_run.node.add_previous_sibling(r_node)
        @runs.insert(@runs.index(first_run), run)
      end
      run
    end

    def replace_in_runs(index, length, replacement)
      total_length = 0
      ends = @runs.map { |r| total_length += r.text_length }
      first_index = ends.index { |e| e > index }

      first_run = @runs[first_index]
      index_in_run = index - (first_index == 0 ? 0 : ends[first_index - 1])
      if ends[first_index] >= index + length
        first_run.text = replace_in_text(first_run.text, index_in_run, length, replacement)
        first_run.adjust_for_right_to_left_text
      else
        length_in_run = first_run.text.length - index_in_run
        first_run.text = replace_in_text(first_run.text, index_in_run, length_in_run, replacement[0,length_in_run])
        first_run.adjust_for_right_to_left_text

        last_index = ends.index { |e| e >= index + length }
        remaining_text = length - length_in_run - clear_runs((first_index + 1), (last_index - 1))

        last_run = last_index.nil? ? @runs.last : @runs[last_index]
        last_run.text = replace_in_text(last_run.text, 0, remaining_text, replacement[length_in_run..-1])
        last_run.adjust_for_right_to_left_text
      end
      [ first_run, index_in_run ]
    end

    def replace_in_text(original, index, length, replacement)
      return original if length == 0
      result = index == 0 ? "" : original[0, index]
      result += replacement unless replacement.nil?
      result += original[(index + length)..-1] unless index + length == original.length
      result
    end
    
    def clear_runs(first, last)
      return 0 unless first <= last
      chars_cleared = 0
      @runs[first..last].each do |r|
        chars_cleared += r.text_length
        r.clear_text
      end
      chars_cleared
    end

    def split_after_run(run)
      r_index = @runs.index(run)
      raise ArgumentError.new("Cannot split paragraph on run that is not in paragraph") if r_index.nil?

      next_node = @node.add_next_sibling("<w:p></w:p>")
      next_node = next_node.first if next_node.is_a? Nokogiri::XML::NodeSet
      next_p = Paragraph.new(next_node, @document)
      @document.paragraph_inserted_after(self, next_p)

      if r_index + 1 < @runs.length
        next_p.runs = @runs.slice!(r_index + 1..-1)
        next_p.runs.each do |r|
          next_node << r.node
          r.paragraph = next_p
        end
      end
    end

    def remove_run(run)
      r_index = @runs.index(run)
      raise ArgumentError.new("Cannot remove run from paragraph to which it does not below") if r_index.nil?

      run.node.remove
      runs.delete_at(r_index)
    end
  end
  
  class Run
    attr_accessor :node
    attr_accessor :text_range
    attr_accessor :paragraph
    
    def initialize(r_node, parent_p)
      @node = r_node
      @paragraph = parent_p
      read_text_range
    end

    def replace_with_run_fragment(fragment)
      new_node = @node.add_next_sibling(fragment)
      new_node = new_node.first if new_node.is_a? Nokogiri::XML::NodeSet
      @node.remove
      @node = new_node
      read_text_range
    end

    def replace_with_body_fragments(fragments)
      @paragraph.split_after_run(self) unless @paragraph.runs.last == self
      @paragraph.remove_run(self)

      fragments.reverse.each do |xml|
        @paragraph.node.add_next_sibling(xml)
        @paragraph.node.next_sibling.xpath(".//w:p").each do |p_node|
          p = Paragraph.new(p_node, @paragraph.document)
          @paragraph.document.paragraph_inserted_after(@paragraph, p)
        end
      end
    end

    def read_text_range
      t_node = @node.at_xpath("w:t")
      @text_range = t_node.nil? ? nil : TextRange.new(t_node)
    end

    def text
      @text_range.nil? ? nil : @text_range.text
    end
    
    def text=(text)
      if text.nil?
        @text_range.node.remove unless @text_range.nil?
        @text_range = nil
      elsif @text_range.nil?
        t_node = Nokogiri::XML::Node.new("w:t", @node.document)
        t_node.content = text
        @node.add_child(t_node)
        @text_range = TextRange.new(t_node)
      else
        @text_range.text = text
      end
    end
    
    def text_length
      @text_range.nil? || @text_range.text.nil? ? 0 : @text_range.text.length
    end
    
    def clear_text
      @text_range.text = "" unless @text_range.nil?
    end

    def adjust_for_right_to_left_text
      return if self.text.nil? or self.text.empty?

      # Can include the following if we ever come across ancient Phoenician sailors needing to do search and replace...
      # \p{Cypriot}\p{Kharoshthi}\p{Lydian}\p{Nko}\p{Phoenician}\p{Syriac}\p{Thaana}
      has_rtl = /[\p{Arabic}\p{Hebrew}]/ =~ text
      return unless has_rtl
      has_non_rtl = /[^\s\d\p{Arabic}\p{Hebrew}]/ =~ text
      return if has_non_rtl

      rPr_node = @node.at_xpath("w:rPr")
      if rPr_node.nil?
        rPr_node = @text_range.node.add_previous_sibling(Nokogiri::XML::Element.new("w:rPr", @node.document))
      end

      rtl_node = rPr_node.at_xpath("w:rtl")
      if rtl_node.nil?
        rPr_node.add_child(Nokogiri::XML::Element.new("w:rtl", rPr_node.document))
      end
    end

    def self.create_image_fragment(image_identifier, pixel_width, pixel_height, image_relationship_id)
      fragment = IO.read(File.join(File.dirname(__FILE__), 'content', 'image_fragment.xml'))
      fragment.gsub!("IMAGE_RELATIONSHIP_ID_PLACEHOLDER", image_relationship_id)
      fragment.gsub!("IDENTIFIER_PLACEHOLDER", image_identifier.to_s)
      fragment.gsub!("EXTENT_WIDTH_PLACEHOLDER", (pixel_height * 6000).to_s)
      fragment.gsub!("EXTENT_LENGTH_PLACEHOLDER", (pixel_width * 6000).to_s)
      fragment
    end
  end

  class TextRange
    attr_accessor :node
    
    def initialize(t_node)
      @node = t_node
    end
    
    def text
      @node.text
    end
    
    def text=(text)
      if text.nil? or text.empty?
        @node.remove_attribute("space")
      else
        @node["xml:space"] = "preserve"
      end
      @node.content = text
    end
  end
end
