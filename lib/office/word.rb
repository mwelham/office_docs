require 'office/package'
require 'office/constants'
require 'office/errors'
require 'office/logger'

module Office
  class WordDocument < Package
    def initialize(filename)
      super(filename)

      main_doc_part = get_relationship_target(WORD_MAIN_DOCUMENT_TYPE)
      raise PackageError.new("Word document package '#{@filename}' has no main document part") if main_doc_part.nil?
      @main_doc = MainDocument.new(self, main_doc_part)
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
    
    def plain_text
      text = ""
      @paragraphs.each do |p| 
        p.runs.each { |r| r.text_ranges.each { |t| text << t.text } }
        text << "\n"
      end
      text
    end
    
    def debug_dump
      p_count = 0
      r_count = 0
      t_chars = 0
      @paragraphs.each do |p|
        p_count += 1
        p.runs.each do |r|
          r_count += 1
          r.text_ranges.each { |t| t_chars += t.text.length unless t.text.nil? }
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
  end
  
  class Run
    attr_accessor :node
    attr_accessor :text_ranges
    
    def initialize(r_node)
      @node = r_node
      @text_ranges = []
      r_node.xpath("w:t").each { |t| @text_ranges << TextRange.new(t) }
    end
  end
  
  class TextRange
    attr_accessor :node
    attr_accessor :text
    
    def initialize(t_node)
      @node = t_node
      @text = t_node.text
    end
  end
end
