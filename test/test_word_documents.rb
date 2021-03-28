#encoding: UTF-8

require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'

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
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_save_simple_doc'
      doc = load_simple_doc
      doc.replace_all("pork chop", "radish and tofu salad")
      doc.save(filename)
      assert File.file?(filename)
      assert File.stat(filename).size > 0
      assert !Office::PackageComparer.are_equal?(SIMPLE_TEST_DOC_PATH, filename)
    end
  end

  def test_save_changes
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_save_simple_doc'

      doc = load_simple_doc
      doc.save(filename)
      assert File.file?(filename)
      assert File.stat(filename).size > 0
      assert Office::PackageComparer.are_equal?(SIMPLE_TEST_DOC_PATH, filename)
    end
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
    doc.add_image test_image
    doc.add_sub_heading "Sub-heading"
    doc.add_image test_image
    doc.add_sub_heading "Sub-heading"
    doc.add_paragraph ""
    doc.add_paragraph "end"

    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_image_addition_doc'
      doc.save(filename)

      doc_copy = Office::WordDocument.new(filename)
      assert_equal doc.plain_text, doc_copy.plain_text
      assert_equal doc_copy.plain_text, "Heading\nintro\nSub-heading\nbody\nSub-heading\n\nSub-heading\n\nSub-heading\n\nend\n"

      assert_not_nil doc_copy.get_part("/word/media/image1.jpeg")
      assert_not_nil doc_copy.get_part("/word/media/image2.jpeg")
      assert_nil doc_copy.get_part("/word/media/image3.jpeg")
    end
  end

  def test_image_replacement
    doc = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'image_replacement_test.docx'))
    doc.replace_all("IMAGE", test_image)

    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_image_addition_doc'
      doc.save(filename)

      doc_copy = Office::WordDocument.new(filename)
      assert_equal doc_copy.plain_text, "Header\n\n\n\nABC\n\nDEF\n\nABCDEF\n\n"

      assert_not_nil doc_copy.get_part("/word/media/image1.jpeg")
      assert_not_nil doc_copy.get_part("/word/media/image2.jpeg")
      assert_not_nil doc_copy.get_part("/word/media/image3.jpeg")
      assert_not_nil doc_copy.get_part("/word/media/image4.jpeg")
      assert_nil doc_copy.get_part("/word/media/image5.jpeg")
    end
  end

  def test_complex_search_and_replace
    source = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'complex_replacement_source.docx'))
    source.replace_all("{{BLOCK_1}}", ["So much Sow!", test_image, nil, "Hopefully crispy"])
    source.replace_all("{{BLOCK_2}}", ["Boudin", "bacon", "ham", "hock", "meatball", "salami", "andouille"])

    target = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'complex_replacement_target.docx'))
    assert docs_are_equivalent?(source, target)
  end

  def test_table_search_and_replace
    source = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'table_replacement_source.docx'))
    source.replace_all("{{MY_TABLE}}", { :column_1 => ["Alpha", "One", 1], :column_2 => ["Bravo", "Two", 2], "Column 3" => nil, "Column 4" => [], :column_5 => "Echo"})

    target = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'table_replacement_target.docx'))
    assert docs_are_equivalent?(source, target)
  end

  def test_image_within_table_search_and_replace
    source = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'image_within_table_replacement_source.docx'))
    source.replace_all("{{MY_TABLE}}", { :column_1 => ["Alpha", "One", 1], :column_2 => ["Bravo", test_image, 2], "Column 3" => ["Charlie", nil, 3]})

    target = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'image_within_table_replacement_target.docx'))
    assert docs_are_equivalent?(source, target)
  end

  def test_table_within_table_search_and_replace
    source = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'table_within_table_replacement_source.docx'))
    inner_table = { :one => ["1", "one"], :two => [ "2", "two"]}
    source.replace_all("{{MY_TABLE}}", { :column_1 => ["Alpha", "One"], :column_2 => ["Bravo", inner_table, 2], "Column 3" => ["Charlie"]})

    target = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'table_within_table_replacement_target.docx'))
    assert docs_are_equivalent?(source, target)
  end

  def test_complex_within_table_search_and_replace
    source = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'complex_within_table_replacement_source.docx'))
    source.replace_all("{{MY_TABLE}}", { :column_1 => "Alpha", :column_2 => [["pre", test_image]], "Column 3" => [["Charlie", "post"]]})

    target = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'complex_within_table_replacement_target.docx'))
    assert docs_are_equivalent?(source, target)
  end

  def test_adding_tables
    source = Office::WordDocument.blank_document
    source.add_heading "Heading"
    source.add_paragraph "intro"
    source.add_sub_heading "Sub-heading"
    source.add_table({ :column_1 => ["Alpha", "One", 1], :column_2 => ["Bravo", "Two", 2], :column_3 => ["Charlie", "Three", 3]})
    source.add_sub_heading "Sub-heading"
    source.add_table({ "Column 1" => ["{{PLACEHOLDER_1}}", ""], "Column 2" => ["", "{{PLACEHOLDER_2}}"], "Column 3" => ["{{PLACEHOLDER_1}}", ""]})
    source.add_paragraph "footer"
    source.replace_all("{{PLACEHOLDER_1}}", "Delta Echo Foxtrot Golf")
    source.replace_all("{{PLACEHOLDER_2}}", "Hotel India Juliet Kilo")

    target = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'add_tables_target.docx'))
    assert docs_are_equivalent?(source, target)
  end

  def test_autosizing_table_columns
    source = Office::WordDocument.blank_document

    table_1 = { :column_1 => ["Alpha", "One", 1], :column_2 => ["Bravo", "Two", 2], :column_3 => ["Charlie", "Three", 3]}
    source.add_heading "Autosized table 1"
    source.add_table(table_1)
    source.add_heading "Full Width table 1"
    source.add_table(table_1, { :use_full_width => true })

    table_2 = { "Single Column" => ["First Row", "Second Row", "Third Row", "Fourth Row"] }
    source.add_heading "Autosized table 2"
    source.add_table(table_2)
    source.add_heading "Full Width table 2"
    source.add_table(table_2, { :use_full_width => true })

    table_3 = {
      :one => ["Asparagus"],
      :two => ["Beetroot"],
      :three => ["Cabbage"],
      :four => ["Dates"],
      :five => ["Elderberries"],
      :six => ["Figs"],
      :seven => ["Grapes"],
      :eight => ["Hackberry"],
      :nine => ["Iceberg lettuce"],
      :ten => ["Jalapeno"],
      :eleven => ["Key lime"]
    }
    source.add_heading "Autosized table 3"
    source.add_table(table_3)
    source.add_heading "Full Width table 3"
    source.add_table(table_3, { :use_full_width => true })

    target = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'autosizing_table_columns_target.docx'))
    assert docs_are_equivalent?(source, target)
  end

  def test_replacing_with_rtl_text
    source = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'rtl_replacement_source.docx'))
    source.replace_all("{{Placeholder}}", "אן כתוב בעברית")

    target = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'rtl_replacement_target.docx'))
    assert docs_are_equivalent?(source, target)
  end

  def test_parsing_headers_and_footers
    docx = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'headers_and_footers.docx'))
    assert_equal "The shirt is nearer the body than the coat.\n", docx.main_doc.plain_text

    assert_equal 1, docx.main_doc.headers.count
    assert_equal "He that has no head, needs no hat.\n", docx.main_doc.headers.first.plain_text

    assert_equal 1, docx.main_doc.footers.count
    assert_equal "If you speak the truth,\n\nkeep a foot in the stirrup.\n", docx.main_doc.footers.first.plain_text
  end

  def test_replacing_placeholders_in_headers_and_footers
    source = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'header_footer_replacement_source.docx'))

    source.replace_all("{{header_text_placeholder}}", "Sitting on a cornflake waiting for the van to come")
    source.replace_all("--body_text_placeholder--", "Elementary penguin singing Hare Krishna")
    source.replace_all("+footer_text_placeholder+", "I am the walrus")

    # TODO Images embedded in headers/footers need to be added to the header/footer's own relationship file.
    # They're presently incorrectly being put in the main doc's _rel file.
    #source.replace_all("{{header_image_placeholder}}", test_image)
    #source.replace_all("__body_image_placeholder__", test_image)
    #source.replace_all("please place the image here", test_image)

    target = Office::WordDocument.new(File.join(File.dirname(__FILE__), 'content', 'header_footer_replacement_target.docx'))
    assert docs_are_equivalent?(source, target)
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

  def test_image
    Magick::ImageList.new File.join(File.dirname(__FILE__), 'content', 'test_image.jpg')
  end

  def docs_are_equivalent?(doc1, doc2)
    xml_1 = doc1.main_doc.part.xml
    xml_2 = doc2.main_doc.part.xml
    EquivalentXml.equivalent?(xml_1, xml_2, { :element_order => true }) { |n1, n2, result| return false unless result }

    # TODO docs_are_equivalent? : check other doc properties

    true
  end
end
