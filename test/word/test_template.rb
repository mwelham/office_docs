#encoding: UTF-8

require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'

class TemplateTest < Test::Unit::TestCase
  SIMPLE_TEST_DOC_PATH = File.join(File.dirname(__FILE__),'..', 'content', 'simple_replacement_test.docx')
  BROKEN_TEST_DOC_PATH = File.join(File.dirname(__FILE__), '..', 'content', 'broken_replacement_test.docx')

  def test_get_placeholders
    doc = load_simple_doc
    template = Word::Template.new(doc)
    placeholders = template.get_placeholders

    assert placeholders == [{:placeholder=>"{{test_food_1}}",
                             :paragraph_index=>2,
                             :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
                             :end_of_placeholder=>{:run_index=>2, :char_index=>8}},
                            {:placeholder=>"{{test_food_2}",
                             :paragraph_index=>2,
                             :beginning_of_placeholder=>{:run_index=>4, :char_index=>1},
                             :end_of_placeholder=>{:run_index=>4, :char_index=>14}},
                            {:placeholder=>"{{ some.cool_heading }}",
                             :paragraph_index=>3,
                             :beginning_of_placeholder=>{:run_index=>1, :char_index=>8},
                             :end_of_placeholder=>{:run_index=>5, :char_index=>2}},
                            {:placeholder=>"{{ lower1}}",
                             :paragraph_index=>6,
                             :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
                             :end_of_placeholder=>{:run_index=>2, :char_index=>1}},
                            {:placeholder=>"{{ lower2}}",
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
      assert(e.message == "Template invalid - end of placeholder }} missing.")
    end
  end

  def test_template_valid
    doc = Office::WordDocument.new(BROKEN_TEST_DOC_PATH)
    template = Word::Template.new(doc)
    refute template.template_valid?
    assert template.errors = "Template invalid - end of placeholder }} missing."
  end

  private

  def load_simple_doc
    Office::WordDocument.new(SIMPLE_TEST_DOC_PATH)
  end

end
