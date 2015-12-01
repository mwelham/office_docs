require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'

# require all the options
Dir[File.join(File.dirname(__FILE__) + '/if_else_replacers', "**/*.rb")].each do |f|
  require f
end

class ForLoopExpanderTest < Test::Unit::TestCase
  IN_SAME_PARAGRAPH_IF_ELSE = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'if_else', 'test_if_in_same_paragraph.docx')

  include IfElseInParagraphTest
  #include IfElseOverParagraphsTest

  #
  def test_get_placeholders
    doc = Office::WordDocument.new(IN_SAME_PARAGRAPH_IF_ELSE)
    template = Word::Template.new(doc)
    placeholders = template.get_placeholders

    correct_placeholder_info = [{:placeholder_text=>"{% if fields.a %}",
  :paragraph_index=>2,
  :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
  :end_of_placeholder=>{:run_index=>0, :char_index=>16}},
 {:placeholder_text=>"{% endif %}",
  :paragraph_index=>2,
  :beginning_of_placeholder=>{:run_index=>0, :char_index=>32},
  :end_of_placeholder=>{:run_index=>0, :char_index=>42}},
 {:placeholder_text=>"{% if fields.b = “pie” %}",
  :paragraph_index=>4,
  :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
  :end_of_placeholder=>{:run_index=>0, :char_index=>24}},
 {:placeholder_text=>"{% endif %}",
  :paragraph_index=>4,
  :beginning_of_placeholder=>{:run_index=>0, :char_index=>34},
  :end_of_placeholder=>{:run_index=>0, :char_index=>44}},
 {:placeholder_text=>"{% if fields.b != “pie” %}",
  :paragraph_index=>6,
  :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
  :end_of_placeholder=>{:run_index=>0, :char_index=>25}},
 {:placeholder_text=>"{% endif %}",
  :paragraph_index=>6,
  :beginning_of_placeholder=>{:run_index=>0, :char_index=>39},
  :end_of_placeholder=>{:run_index=>0, :char_index=>49}},
 {:placeholder_text=>"{% if fields.c == “lol” %}",
  :paragraph_index=>8,
  :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
  :end_of_placeholder=>{:run_index=>1, :char_index=>9}},
 {:placeholder_text=>"{% endif %}",
  :paragraph_index=>8,
  :beginning_of_placeholder=>{:run_index=>1, :char_index=>22},
  :end_of_placeholder=>{:run_index=>1, :char_index=>32}}]

    placeholders.each do |p|
      p.delete(:paragraph_object)
    end

    assert correct_placeholder_info == placeholders
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
