require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'

class ForLoopExpanderTest < Test::Unit::TestCase
  IN_SAME_PARAGRAPH_FOR_LOOP = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'for_loops', 'in_same_paragraph_for_loop_test.docx')
  IN_DIFFERENT_PARAGRAPH_FOR_LOOP = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'for_loops', 'in_different_paragraph_for_loop_test.docx')

  ##
  # def test_get_placeholders
  #   doc = Office::WordDocument.new(IN_SAME_PARAGRAPH_FOR_LOOP)
  #   template = Word::Template.new(doc)
  #   placeholders = template.get_placeholders

  #   correct_placeholder_info = [{:placeholder_text=>"{% foreach field in fields.Group %}",
  #                                 :paragraph_index=>2,
  #                                 :beginning_of_placeholder=>{:run_index=>0, :char_index=>0},
  #                                 :end_of_placeholder=>{:run_index=>4, :char_index=>2}},
  #                               {:placeholder_text=>"{% endeach %}",
  #                                 :paragraph_index=>2,
  #                                 :beginning_of_placeholder=>{:run_index=>4, :char_index=>13},
  #                                 :end_of_placeholder=>{:run_index=>6, :char_index=>2}}]

  #   correct_placeholder_info.each do |placeholder_info|
  #     target = placeholders.find{|p| p[:placeholder_text] == placeholder_info[:placeholder_text]}
  #     assert target[:paragraph_index] == placeholder_info[:paragraph_index]
  #     assert target[:beginning_of_placeholder] == placeholder_info[:beginning_of_placeholder]
  #     assert target[:end_of_placeholder] == placeholder_info[:end_of_placeholder]
  #   end
  # end

  ##
  # def test_loop_in_same_paragraph
  #   file = File.new('test_save_simple_doc.docx', 'w')
  #   file.close
  #   filename = file.path

  #   doc = Office::WordDocument.new(IN_SAME_PARAGRAPH_FOR_LOOP)
  #   template = Word::Template.new(doc)
  #   template.render(
  #     {'fields' =>
  #       {'Group' => [
  #         {'Q' => 'a'},
  #         {'Q' => 'b'},
  #         {'Q' => 'c'}
  #       ]
  #     }
  #   })
  #   template.word_document.save(filename)

  #   assert File.file?(filename)
  #   assert File.stat(filename).size > 0

  #   correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'for_loops', 'correct_render', 'in_same_paragraph_for_loop_test.docx'))
  #   our_render = Office::WordDocument.new(filename)
  #   assert docs_are_equivalent?(correct, our_render)

  #   File.delete(filename)
  # end

  ##
  # def test_loop_in_different_paragraph_with_blank_group
  #   file = File.new('test_save_simple_doc.docx', 'w')
  #   file.close
  #   filename = file.path

  #   doc = Office::WordDocument.new(IN_DIFFERENT_PARAGRAPH_FOR_LOOP)
  #   template = Word::Template.new(doc)
  #   template.render(
  #     {'fields' =>
  #       {'Group' => []
  #     }
  #   })
  #   template.word_document.save(filename)

  #   assert File.file?(filename)
  #   assert File.stat(filename).size > 0

  #   correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'for_loops', 'correct_render', 'in_different_paragraph_blank_group.docx'))
  #   our_render = Office::WordDocument.new(filename)
  #   assert docs_are_equivalent?(correct, our_render)

  #   File.delete(filename)
  # end

  def test_loop_in_different_paragraph
    file = File.new('test_save_simple_doc.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(IN_DIFFERENT_PARAGRAPH_FOR_LOOP)
    template = Word::Template.new(doc)
    template.render(
      {'fields' =>
        {'Group' => [
          {'Q' => 'a'},
          {'Q' => 'b'},
          {'Q' => 'c'}
        ]
      }
    })
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    binding.pry

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'for_loops', 'correct_render', 'in_different_paragraph_for_loop_test.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
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
