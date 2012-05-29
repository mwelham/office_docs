require 'test/unit'
require 'office_docs'

class WordDocParsingTest < Test::Unit::TestCase
  def test_parse_simple_doc
    file_path = File.join(File.dirname(__FILE__), 'content', 'simple_test.docx')
    doc = Office::WordDocument.new(file_path)
    doc.debug_dump
  end
end
