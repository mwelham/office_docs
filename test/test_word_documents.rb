require 'test/unit'
require 'date'
require 'office_docs'

class WordDocumentsTest < Test::Unit::TestCase
  SIMPLE_TEST_DOC_PATH = File.join(File.dirname(__FILE__), 'content', 'simple_test.docx')
  COMPLEX_TEST_DOC_PATH = File.join(File.dirname(__FILE__), 'content', 'complex_test.docx')

  def test_parse_simple_doc
    doc = load_simple_doc
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
    
    doc = load_simple_doc
    doc.replace_all("pork chop", "radish and tofu salad")
    doc.save(filename)
    assert File.file?(filename)
    assert File.stat(filename).size > 0
    assert !Office::PackageComparer.are_equal?(SIMPLE_TEST_DOC_PATH, filename)
    
    file.delete
  end
  
  def test_save_changes
    file = Tempfile.new('test_save_simple_doc')
    file.close
    filename = file.path
    
    doc = load_simple_doc
    doc.save(filename)
    assert File.file?(filename)
    assert File.stat(filename).size > 0
    assert Office::PackageComparer.are_equal?(SIMPLE_TEST_DOC_PATH, filename)
    
    file.delete
  end

  def test_blank_document
    assert_equal Office::WordDocument.blank_document.plain_text, ""
  end

  def test_build_document
    doc = Office::WordDocument.blank_document
    doc.add_heading "Heading"
    doc.add_paragraph "intro"
    doc.add_sub_heading "Sub-heading"
    doc.add_paragraph "body"
    doc.add_paragraph ""
    doc.add_paragraph "end"
    assert_equal doc.plain_text, "Heading\nintro\nSub-heading\nbody\n\nend\n"
  end

  def test_from_data
    doc_1 = nil
    File.open(SIMPLE_TEST_DOC_PATH) { |f| doc_1 = Office::WordDocument.from_data(f.read) }
    doc_2 = load_simple_doc
    assert_equal doc_1.plain_text, doc_2.plain_text
  end

  def test_to_data
    data = load_simple_doc.to_data
    assert !data.nil?
    assert data.length > 0

    doc_1 = Office::WordDocument.from_data(data)
    doc_2 = load_simple_doc
    assert_equal doc_1.plain_text, doc_2.plain_text
  end

  def test_complex_parsing
    doc = Office::WordDocument.new(COMPLEX_TEST_DOC_PATH)
    assert doc.plain_text.include?("Presiding Peasant: Dennis")
    assert doc.plain_text.include?("Assessment Depot: Swampy Castle (might be sinking)")
    replace_and_check(doc, "Swampy Castle (might be sinking)", "Farcical Aquatic Ceremony")
  end

  def test_image_addition
    doc = Office::WordDocument.blank_document
    doc.add_heading "Heading"
    doc.add_paragraph "intro"
    doc.add_sub_heading "Sub-heading"
    doc.add_paragraph "body"
    doc.add_sub_heading "Sub-heading"
    doc.add_image Magick::ImageList.new File.join(File.dirname(__FILE__), 'content', 'test_image.jpg')
    doc.add_sub_heading "Sub-heading"
    doc.add_image Magick::ImageList.new File.join(File.dirname(__FILE__), 'content', 'test_image.jpg')
    doc.add_sub_heading "Sub-heading"
    doc.add_paragraph ""
    doc.add_paragraph "end"

    file = Tempfile.new('test_image_addition_doc')
    file.close
    filename = file.path
    doc.save(filename)

    doc_copy = Office::WordDocument.new(filename)
    assert_equal doc.plain_text, doc_copy.plain_text
    assert_equal doc_copy.plain_text, "Heading\nintro\nSub-heading\nbody\nSub-heading\n\nSub-heading\n\nSub-heading\n\nend\n"

    assert_not_nil doc_copy.get_part("/word/media/image1.jpeg")
    assert_not_nil doc_copy.get_part("/word/media/image2.jpeg")
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
