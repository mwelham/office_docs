require 'office/package'
require 'office/constants'
require 'office/errors'
require 'office/logger'

module Office
  class WordDocument < Package
    attr_accessor :main_doc
    
    def initialize(filename)
      super(filename)

      main_doc_part = get_relationship_target(WORD_MAIN_DOCUMENT_TYPE)
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

    def add_image(image) # image must be an Magick::Image
      p = @main_doc.add_paragraph
      p.add_run_with_fragment(create_image_fragment(image))
      p
    end

    def plain_text
      @main_doc.plain_text
    end

    # The type of 'replacement' determines what replaces the source text:
    #   Image  - an image (Magick::Image or Magick::ImageList)
    #   Hash   - a table, keys being column headings, and each value an array of column data
    #   Array  - a sequence of these replacement types all of which will be inserted
    #   String - simple text replacement
    def replace_all(source_text, replacement)
      if replacement.kind_of? String
        # Special case for simple text, so we can preserve the style of source in the replacement
        @main_doc.replace_all_with_text(source_text, replacement)
        return
      end

      runs = @main_doc.replace_all_with_empty_runs(source_text)
      runs.each { |r| r.replace_with_fragment(create_fragment(replacement)) }
    end

    def create_fragment(item)
      case
      when (item.is_a?(Magick::Image) or item.is_a?(Magick::ImageList))
        create_image_fragment(item)
      when item.is_a?(Hash)
        create_table_fragment(item)
      when item.is_a?(Array)
        create_mulitple_fragments(item)
      else
        create_text_fragment(item.nil? ? "" : item.to_s)
      end
    end

    def create_image_fragment(image)
      prefix = ["", @main_doc.part.path_components, "media", "image"].flatten.join('/')
      identifier = unused_part_identifier(prefix)
      extension = "#{image.format}".downcase

      part = add_part("#{prefix}#{identifier}.#{extension}", StringIO.new(image.to_blob), image.mime_type)
      relationship_id = @main_doc.part.add_relationship(part, IMAGE_RELATIONSHIP_TYPE)

      Run.create_image_fragment(identifier, image.columns, image.rows, relationship_id)
    end

    def create_table_fragment(hash)
      # TODO WordDocument.create_table_fragment
      create_text_fragment("(tables are not yet implemented)")
    end

    def create_mulitple_fragments(array)
      array.inject("") { |fragments, item| fragments + create_fragment(item) }
    end

    def create_text_fragment(text)
      "<w:r><w:t>#{Nokogiri::XML::Document.new.encode_special_chars(text)}</w:t></w:r>"
    end

    def debug_dump
      super
      @main_doc.debug_dump
      #Logger.debug_dump_xml("Word Main Document", @main_doc.part.xml)
    end
  end
  
  class MainDocument
    attr_accessor :part
    attr_accessor :body_node
    attr_accessor :paragraphs
    
    def initialize(word_doc, part)
      @parent = word_doc
      @part = part
      parse_xml
    end
    
    def parse_xml
      xml_doc = @part.xml
      @body_node = xml_doc.at_xpath("/w:document/w:body")
      raise PackageError.new("Word document '#{@filename}' is missing main document body") if body_node.nil?

      @paragraphs = []
      body_node.xpath(".//w:p").each { |p| @paragraphs << Paragraph.new(p) }
    end

    def add_paragraph
      p = @body_node.document.create_element("p")
      p_node = @paragraphs.empty? ? @body_node.add_child(p) : @paragraphs.last.node.add_next_sibling(p)
      @paragraphs << Paragraph.new(p_node)
      @paragraphs.last
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

    def debug_dump
      p_count = 0
      r_count = 0
      t_chars = 0
      @paragraphs.each do |p|
        p_count += 1
        p.runs.each do |r|
          r_count += 1
          t_chars += r.text_length
        end
      end
      Logger.debug_dump "Main Document Stats"
      Logger.debug_dump "  paragraphs  : #{p_count}"
      Logger.debug_dump "  runs        : #{r_count}"
      Logger.debug_dump "  text length : #{t_chars}"
      Logger.debug_dump ""

      Logger.debug_dump "Main Document Plain Text"
      Logger.debug_dump ">>>"
      Logger.debug_dump plain_text
      Logger.debug_dump "<<<"
      Logger.debug_dump ""
    end
  end
  
  class Paragraph
    attr_accessor :node
    attr_accessor :runs
    
    def initialize(p_node)
      @node = p_node
      @runs = []
      p_node.xpath("w:r").each { |r| @runs << Run.new(r) }
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

      r = Run.new(r_node)
      @runs << r
      r
    end

    def populate_r_node(r_node, text)
      t_node = r_node.add_child(@node.document.create_element("t"))
      t_node["xml:space"] = "preserve"
      t_node.content = text
    end

    def add_run_with_fragment(fragment)
      r = Run.new(@node.add_child(fragment))
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
      run = Run.new(r_node)
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
        @runs.insert(@runs.index(first_run), Run.new(preceding_r_node))

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
      else
        length_in_run = first_run.text.length - index_in_run
        first_run.text = replace_in_text(first_run.text, index_in_run, length_in_run, replacement[0,length_in_run])

        last_index = ends.index { |e| e >= index + length }
        remaining_text = length - length_in_run - clear_runs((first_index + 1), (last_index - 1))

        last_run = last_index.nil? ? @runs.last : @runs[last_index]
        last_run.text = replace_in_text(last_run.text, 0, remaining_text, replacement[length_in_run..-1])
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
  end
  
  class Run
    attr_accessor :node
    attr_accessor :text_range
    
    def initialize(r_node)
      @node = r_node
      read_text_range
    end

    def replace_with_fragment(fragment)
      new_node = @node.add_next_sibling(fragment)
      @node.remove
      @node = new_node
      read_text_range
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
      @node.content = text
    end
  end
end
