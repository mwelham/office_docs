#encoding: UTF-8

require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'

class TemplateTest < Test::Unit::TestCase
  SIMPLE_TEST_DOC_PATH = File.join(File.dirname(__FILE__),'..', 'content', 'template', 'simple_replacement_test.docx')
  BROKEN_TEST_DOC_PATH = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'broken_replacement_test.docx')
  BIG_TEST_DOC_PATH = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'really_big_template.docx')
  MENICON_DOC_PATH = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'menicon_template.docx')

  TEMPLATE_ALL_OPTIONS = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'Template_Test_Form.docx')

  def test_get_placeholders
    doc = Office::WordDocument.new(SIMPLE_TEST_DOC_PATH)
    template = Word::Template.new(doc)
    placeholders = template.get_placeholders

    assert placeholders == [{:placeholder_text=>"{{test_food_1}}",
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
    file = File.new('test_save_simple_doc.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(SIMPLE_TEST_DOC_PATH)
    template = Word::Template.new(doc)
    template.render({})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    File.delete(filename)
  end

  def test_render_big_template
    file = File.new('test_save_simple_doc.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(BIG_TEST_DOC_PATH)
    template = Word::Template.new(doc)
    template.render({})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    File.delete(filename)
  end

  def test_render_menicon_template
    file = File.new('test_menicon_template.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(MENICON_DOC_PATH)
    template = Word::Template.new(doc)
    template.render({})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    File.delete(filename)
  end

  def test_template_all_options
    file = File.new('test_template_all_options.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(TEMPLATE_ALL_OPTIONS)
    template = Word::Template.new(doc)
    template.render(test_template_all_options_test_data)
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'correct_render', 'test_template_all_options.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
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
