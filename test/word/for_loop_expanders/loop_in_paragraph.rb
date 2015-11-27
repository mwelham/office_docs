module LoopInParagraphTest
  IN_SAME_PARAGRAPH_FOR_LOOP = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'in_same_paragraph_for_loop_test.docx')

  #SAME PARAGRAPH LOOPS - RUN LOOPS
  #
  #
  #
  def test_loop_in_same_paragraph
    file = File.new('test_save_simple_doc.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(IN_SAME_PARAGRAPH_FOR_LOOP)
    template = Word::Template.new(doc)
    template.render(
      {'fields' =>
        {'Group' => [
          {'Q' => 'a'},
          {'Q' => 'b'},
          {'Q' => 'c'}
        ]
      }
    }, {do_not_render: true})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'in_same_paragraph_for_loop_test.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end
end
