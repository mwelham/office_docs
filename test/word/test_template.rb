#encoding: UTF-8

require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'

class TemplateTest < Test::Unit::TestCase
  SIMPLE_TEST_DOC_PATH = File.join(File.dirname(__FILE__),'..', 'content', 'template', 'placeholders', 'simple_replacement_test.docx')
  BROKEN_TEST_DOC_PATH = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'broken_replacement_test.docx')
  BIG_TEST_DOC_PATH = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'really_big_template.docx')
  MENICON_DOC_PATH = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'menicon_template.docx')
  NESTED_PLACEHOLDERS = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'nested_placeholders.docx')
  TEST_QUOTE_MARKS_DOC_PATH = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'quote_mark_in_template_options.docx')

  TEMPLATE_ALL_OPTIONS = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'Template_Test_Form.docx')
  TEST_ARABIC_DATE_TIME = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'test_arabic_date_time.docx')

  def test_get_placeholders
    doc = Office::WordDocument.new(SIMPLE_TEST_DOC_PATH)
    template = Word::Template.new(doc)
    placeholders = template.get_placeholders

    correct_placeholder_info = [{:placeholder_text=>"{{test_food_1}}",
                             :paragraph_index=>2,
                             :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
                             :end_of_placeholder=>{:run_index=>2, :char_index=>8}},
                            {:placeholder_text=>"{{test_food_2}}",
                             :paragraph_index=>2,
                             :beginning_of_placeholder=>{:run_index=>4, :char_index=>1},
                             :end_of_placeholder=>{:run_index=>4, :char_index=>15}},
                            {:placeholder_text=>"{{ some.cool_heading }}",
                             :paragraph_index=>3,
                             :beginning_of_placeholder=>{:run_index=>1, :char_index=>8},
                             :end_of_placeholder=>{:run_index=>5, :char_index=>2}},
                            {:placeholder_text=>"{{ lower1}}",
                             :paragraph_index=>6,
                             :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
                             :end_of_placeholder=>{:run_index=>2, :char_index=>1}},
                            {:placeholder_text=>"{{ lower2}}",
                             :paragraph_index=>8,
                             :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
                             :end_of_placeholder=>{:run_index=>2, :char_index=>1}}]

    correct_placeholder_info.each do |placeholder_info|
      target = placeholders.find{|p| p[:placeholder_text] == placeholder_info[:placeholder_text]}
      assert target[:paragraph_index] == placeholder_info[:paragraph_index]
      assert target[:beginning_of_placeholder] == placeholder_info[:beginning_of_placeholder]
      assert target[:end_of_placeholder] == placeholder_info[:end_of_placeholder]
    end
  end

  def test_get_nested_placeholders
    doc = Office::WordDocument.new(NESTED_PLACEHOLDERS)
    template = Word::Template.new(doc)

    good_placeholders=  [{:placeholder_text=>"{{fields.CLIENT.Client_First_Name}}",
      :paragraph_index=>2,
      :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
      :end_of_placeholder=>{:run_index=>0, :char_index=>34}},
     {:placeholder_text=>"{{fields.CLIENT.Client_Last_Name}}",
      :paragraph_index=>2,
      :beginning_of_placeholder=>{:run_index=>0, :char_index=>37},
      :end_of_placeholder=>{:run_index=>0, :char_index=>70}},
     {:placeholder_text=>"{{fields.CLIENT.Client_Add1}}",
      :paragraph_index=>3,
      :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
      :end_of_placeholder=>{:run_index=>0, :char_index=>28}},
     {:placeholder_text=>"{% if fields.CLIENT.Client_Add2 %}",
      :paragraph_index=>4,
      :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
      :end_of_placeholder=>{:run_index=>0, :char_index=>33}},
     {:placeholder_text=>"{{fields.CLIENT.Client_Add2}}",
      :paragraph_index=>4,
      :beginning_of_placeholder=>{:run_index=>0, :char_index=>34},
      :end_of_placeholder=>{:run_index=>0, :char_index=>62}},
     {:placeholder_text=>"{% endif %}",
      :paragraph_index=>4,
      :beginning_of_placeholder=>{:run_index=>0, :char_index=>63},
      :end_of_placeholder=>{:run_index=>0, :char_index=>73}},
     {:placeholder_text=>"{{fields.CLIENT.Client_City}}",
      :paragraph_index=>5,
      :beginning_of_placeholder=>{:run_index=>0, :char_index=>2},
      :end_of_placeholder=>{:run_index=>0, :char_index=>30}},
     {:placeholder_text=>"{{fields.CLIENT.Client_State}}",
      :paragraph_index=>5,
      :beginning_of_placeholder=>{:run_index=>0, :char_index=>33},
      :end_of_placeholder=>{:run_index=>0, :char_index=>62}},
     {:placeholder_text=>"{{fields.CLIENT.Client_Zip}}",
      :paragraph_index=>5,
      :beginning_of_placeholder=>{:run_index=>0, :char_index=>64},
      :end_of_placeholder=>{:run_index=>0, :char_index=>91}}]

    placeholders = template.get_placeholders
    placeholders.map {|placeholder| placeholder.delete(:paragraph_object)}
    assert placeholders == good_placeholders

  end

  def test_broken_template
    doc = Office::WordDocument.new(BROKEN_TEST_DOC_PATH)
    template = Word::Template.new(doc)
    begin
      placeholders = template.get_placeholders
    rescue => e
      assert(e.message == "Template invalid - end of placeholder }} missing for \"{{test_food_2 pancetta pork chuck jowl pig.\".")
    end
  end

  def test_template_valid
    doc = Office::WordDocument.new(BROKEN_TEST_DOC_PATH)
    template = Word::Template.new(doc)
    refute template.template_valid?
    assert template.errors = "Template invalid - end of placeholder }} missing."
  end

  def test_render
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_save_simple_doc.docx'

      doc = Office::WordDocument.new(SIMPLE_TEST_DOC_PATH)
      template = Word::Template.new(doc)
      template.render({})
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0
    end
  end

  def test_render_big_template
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_save_simple_doc.docx'

      doc = Office::WordDocument.new(BIG_TEST_DOC_PATH)
      template = Word::Template.new(doc)
      template.render({})
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0
    end
  end

  def test_render_menicon_template
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_menicon_template.docx'

      doc = Office::WordDocument.new(MENICON_DOC_PATH)
      template = Word::Template.new(doc)
      template.render({})
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0
    end
  end

  def test_template_all_options
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_template_all_options.docx'

      doc = Office::WordDocument.new(TEMPLATE_ALL_OPTIONS)
      template = Word::Template.new(doc)
      template.render(test_template_all_options_test_data)
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'correct_render', 'test_template_all_options.docx'))
      our_render = Office::WordDocument.new(filename)
      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def test_quote_markes_in_options
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_quote_marks_template.docx'

      doc = Office::WordDocument.new(TEST_QUOTE_MARKS_DOC_PATH)
      template = Word::Template.new(doc)
      template.render(test_template_all_options_test_data)
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'correct_render', 'test_quote_marks_template.docx'))
      our_render = Office::WordDocument.new(filename)
      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def test_quote_markes_in_options
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_quote_marks_template.docx'

      doc = Office::WordDocument.new(TEST_QUOTE_MARKS_DOC_PATH)
      template = Word::Template.new(doc)
      template.render(test_template_all_options_test_data)
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'correct_render', 'test_quote_marks_template.docx'))
      our_render = Office::WordDocument.new(filename)
      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def test_parse_arabic_datetime
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_arabic_date_time.docx'

      doc = Office::WordDocument.new(TEST_ARABIC_DATE_TIME)
      template = Word::Template.new(doc)
      template.render({
          'fields' => {
            'a' => '١٩٨٧/٠٦/١٩ ٢٠:٠٠:٠٠'
          }
        })
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'placeholders', 'correct_render', 'test_arabic_date_time.docx'))
      our_render = Office::WordDocument.new(filename)

      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def test_image_resizing
    doc = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'image_resize_noresample_test.docx'))
    doc.render_template({"IMAGE"=> test_image})

    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_image_resize_doc'
      doc.save(filename)

      doc_copy = Office::WordDocument.new(filename)
      assert_equal doc_copy.plain_text, "Header\n\nImage resize and noresample test\n\n\n\n\n\n\n\n"

      # Normal image (640x480 px)
      default_image = doc_copy.get_part("/word/media/image1.jpeg")
      assert_not_nil default_image
      assert_equal default_image.image.columns, 640
      assert_equal default_image.image.rows, 480

      # Resized and resampled (100x75 px)
      resized_image = doc_copy.get_part("/word/media/image2.jpeg")
      assert_not_nil resized_image
      assert_equal resized_image.image.columns, 100
      assert_equal resized_image.image.rows, 75

      # Resized, not resampled (640x480 px)
      noresample_image = doc_copy.get_part("/word/media/image3.jpeg")
      assert_not_nil noresample_image
      assert_equal noresample_image.image.columns, 640
      assert_equal noresample_image.image.rows, 480

      assert_nil doc_copy.get_part("/word/media/image4.jpeg")
    end
  end



  private

  def test_template_all_options_test_data
    {"device_id"=>"iPhone_C057D24A-F40F-4182-8197-570EFAB4232A",
     "username"=>"Matts gain",
     "submitted_at"=>"2015-08-21 15:11:08 +02:00",
     "received_at"=>"2015-08-21 13:11:09 +00:00",
     "submission_id"=>"110",
     "device_submission_identifier"=>"F5E1EC21-F18B-4321-A77B-D1C65A831D48",
     "form_name"=>"Template Test Form",
     "form_namespace"=>"http://www.devicemagic.com/xforms/c5b2cd00-2a26-0133-4595-14109fd23119",
     "device"=>{"Test"=>nil, "Test_Attribute"=>nil, "Testing"=>"2", "Weo"=>nil, "zxczxc"=>nil},
     "fields"=>
      {"Free_Text_Question"=>"I\nAm\nCool\n",
       "Untitled_Question"=>"Here",
       "Decimal_Question"=>"1.8",
       "Select_Question"=>"2,4",
       "Group"=>
        [{"Free_Text_Question"=>"Q",
          "Image_Question"=> test_image,
          "Location_Question"=>[ test_image, "lat=-33.863495, long=18.641860, alt=166.692398, hAccuracy=104.580297, vAccuracy=10.000000, timestamp=2015-08-21T13:09:54Z"],
          "Group"=>[{"Free_Text_Question"=>"1"}, {"Free_Text_Question"=>"2"}]},
         {"Free_Text_Question"=>"Q2",
          "Image_Question"=> test_image,
          "Location_Question"=>[ test_image, "lat=-33.863396, long=18.641903, alt=167.221832, hAccuracy=65.000000, vAccuracy=10.000000, timestamp=2015-08-21T13:10:57Z"],
          "Group"=>[{"Free_Text_Question"=>"3"}, {"Free_Text_Question"=>"4"}]}],
       "Image_Question"=> test_image,
       "Location_Question"=>[ test_image, "lat=-33.863396, long=18.641903, alt=167.221832, hAccuracy=65.000000, vAccuracy=10.000000, timestamp=2015-08-21T13:11:04Z"]}}
  end

  def test_image
    Magick::ImageList.new File.join(File.dirname(__FILE__), '..', 'content', 'test_image.jpg')
  end

  def docs_are_equivalent?(doc1, doc2)
    xml_1 = doc1.main_doc.part.xml
    xml_2 = doc2.main_doc.part.xml
    EquivalentXml.equivalent?(xml_1, xml_2, { :element_order => true }) { |n1, n2, result| return false unless result }

    # TODO docs_are_equivalent? : check other doc properties

    true
  end

end
