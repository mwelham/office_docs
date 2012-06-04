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
      WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'blank.docx'))
    end
    
    def add_heading(text)
      p = @main_doc.add_paragraph
      p.add_style("Heading1")
      p.add_run(text)
      p
    end
    
    def add_sub_heading(text)
      p = @main_doc.add_paragraph
      p.add_style("Heading2")
      p.add_run(text)
      p
    end
    
    def add_paragraph(text)
      p = @main_doc.add_paragraph
      p.add_run(text)
      p
    end
    
    def plain_text
      @main_doc.plain_text
    end
    
    def replace_all(source, replacement)
      @main_doc.replace_all(source, replacement)
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
      body_node.xpath("w:p").each { |p| @paragraphs << Paragraph.new(p) }
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
        p.runs.each { |r| text << r.text }
        text << "\n"
      end
      text
    end

    def replace_all(source, replacement)
      @paragraphs.each { |p| p.replace_all(source, replacement) }
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

    def add_run(text)
      r_node = @node.add_child(@node.document.create_element("r"))
      t_node = r_node.add_child(@node.document.create_element("t"))
      t_node["xml:space"] = "preserve"
      t_node.content = text

      r = Run.new(r_node)
      @runs << r
      r
    end

    def replace_all(source, replacement)
      return if source.nil? or source.empty?
      replacement = "" if replacement.nil?
      
      text = @runs.inject("") { |t, run| t + run.text }
      until (i = text.index(source, i.nil? ? 0 : i)).nil?
        replace_in_runs(i, source.length, replacement)
        text = replace_in_text(text, i, source.length, replacement)
        i += replacement.length
      end
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
      t_node = r_node.at_xpath("w:t")
      @text_range = TextRange.new(t_node) unless t_node.nil?
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
