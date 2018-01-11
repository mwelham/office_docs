require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'
require 'helpers/template_test_helper'

# require all the options
Dir[File.join(File.dirname(__FILE__) + '/for_loop_expanders', "**/*.rb")].each do |f|
  require f
end

class ForLoopExpanderTest < Test::Unit::TestCase
  IN_SAME_PARAGRAPH_FOR_LOOP = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'for_loops', 'in_same_paragraph_for_loop_test.docx')
  MISSING_END_FOR = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'for_loops', 'missing_end_for.docx')

  include TemplateTestHelper
  include LoopInParagraphTest
  include LoopOverParagraphsTest
  include LoopTableRowTest

  #
  def test_get_placeholders
    doc = Office::WordDocument.new(IN_SAME_PARAGRAPH_FOR_LOOP)
    template = Word::Template.new(doc)
    placeholders = template.get_placeholders

    correct_placeholder_info = [{:placeholder_text=>"{% for field in fields.Group %}",
                                :paragraph_index=>2,
                                :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
                                :end_of_placeholder=>{:run_index=>0, :char_index=>30}},
                               {:placeholder_text=>"{{ field.name }}",
                                :paragraph_index=>2,
                                :beginning_of_placeholder=>{:run_index=>1, :char_index=>1},
                                :end_of_placeholder=>{:run_index=>1, :char_index=>16}},
                               {:placeholder_text=>"{% endfor %}",
                                :paragraph_index=>2,
                                :beginning_of_placeholder=>{:run_index=>2, :char_index=>7},
                                :end_of_placeholder=>{:run_index=>4, :char_index=>2}},
                               {:placeholder_text=>"{% for field in fields.Group %}",
                                :paragraph_index=>2,
                                :beginning_of_placeholder=>{:run_index=>5, :char_index=>16},
                                :end_of_placeholder=>{:run_index=>7, :char_index=>17}},
                               {:placeholder_text=>"{{ field.age }}",
                                :paragraph_index=>2,
                                :beginning_of_placeholder=>{:run_index=>8, :char_index=>1},
                                :end_of_placeholder=>{:run_index=>8, :char_index=>15}},
                               {:placeholder_text=>"{% endfor %}",
                                :paragraph_index=>2,
                                :beginning_of_placeholder=>{:run_index=>9, :char_index=>16},
                                :end_of_placeholder=>{:run_index=>11, :char_index=>2}}]



    placeholders.each do |p|
      p.delete(:paragraph_object)
    end

    assert correct_placeholder_info == placeholders
  end

  def test_missing_end_for
    doc = Office::WordDocument.new(MISSING_END_FOR)
    template = Word::Template.new(doc)
    err = assert_raises ::RuntimeError do
      template.render({})
    end
    assert err.message == "Missing endfor for 'for each' placeholder: {% for field in fields.Group %}"
  end

  private

  def docs_are_equivalent?(doc1, doc2)
    xml_1 = doc1.main_doc.part.xml
    xml_2 = doc2.main_doc.part.xml
    EquivalentXml.equivalent?(xml_1, xml_2, { :element_order => true }) { |n1, n2, result| return false unless result }

    # TODO docs_are_equivalent? : check other doc properties

    true
  end
end
