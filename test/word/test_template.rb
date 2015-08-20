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

end
