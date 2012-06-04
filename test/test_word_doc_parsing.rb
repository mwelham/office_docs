require 'test/unit'
require 'date'
require 'office_docs'

class WordDocParsingTest < Test::Unit::TestCase
  SIMPLE_TEST_DOC_PATH = File.join(File.dirname(__FILE__), 'content', 'simple_test.docx')

  def test_parse_simple_doc
    doc = load_simple_doc
    #doc.debug_dump
  end

  def test_replace
    replace_and_check(load_simple_doc, "pork", "lettuce")
    replace_and_check(load_simple_doc, "lettuce", "pork")
    replace_and_check(load_simple_doc, "pork", "pork")
    replace_and_check(load_simple_doc, "Short ribs meatball pork chop sausage, ham hock biltong cow", "..")
    replace_and_check(load_simple_doc, "Simple Test Document", "")
    #stress_test_replace(SIMPLE_TEST_DOC_PATH)
  end

  def test_save_simple_doc
    file = Tempfile.new('test_save_simple_doc')
    file.close
    filename = file.path
    #filename = "/Users/mwelham/Temp/test_save_simple_doc.docx"
    
    doc = load_simple_doc
    doc.save(filename)
    assert File.file?(filename)
    assert File.stat(filename).size > 0
    
    file.delete
  end
  
  private

  def load_simple_doc
    Office::WordDocument.new(SIMPLE_TEST_DOC_PATH)
  end

  def replace_and_check(doc, source, replacement)
    original = doc.plain_text
    doc.replace_all(source, replacement)
    assert_equal original.gsub(source, replacement), doc.plain_text
  end

  def stress_test_replace(doc_path)
    1000.times do
      doc = Office::WordDocument.new(doc_path)
      replace_and_check(doc, random_substring(doc.plain_text), random_text)
    end
  end

  def random_substring(text)
    substring = "\n"
    while substring.include? "\n" do
      start = Random::rand(text.length - 1)
      length = 1 + Random::rand(text.length - start - 1)
      substring = text[start, length]
    end
    substring
  end

  def random_text
    text = ""
    Random::rand(20).times { text <<= 'a' }
    text
  end
end
