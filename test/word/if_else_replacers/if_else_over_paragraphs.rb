module IfElseOverParagraphsTest
  OVER_PARAGRAPHS_IF_ELSE = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'test_if_over_paragraphs.docx')

  #SAME PARAGRAPH LOOPS - RUN LOOPS
  #
  #
  #
  def test_if_else_over_paragraphs
    file = File.new('test_if_over_paragraphs.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(OVER_PARAGRAPHS_IF_ELSE)
    template = Word::Template.new(doc)
    template.render(
      {'fields' =>
        {
          'a' => '',
          'b' => 'haha',
          'c' => 'lol'
        }
      }, {do_not_render: true})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'correct_render', 'test_if_over_paragraphs.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end
end
