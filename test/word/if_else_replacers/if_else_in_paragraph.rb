module IfElseInParagraphTest
  IN_SAME_PARAGRAPH_IF_ELSE = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'test_if_in_same_paragraph.docx')
  IF_ELSE_IN_TABLES = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'if_else_in_tables.docx')
  IN_SAME_PARAGRAPH_NESTED_IF_ELSE = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'nested_if_else_in_same_paragraph.docx')
  IF_WITH_AND = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'test_if_with_and.docx')

  ##SAME PARAGRAPH LOOPS - RUN LOOPS



  def test_if_else_in_same_paragraph
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_if_in_same_paragraph.docx'

      doc = Office::WordDocument.new(IN_SAME_PARAGRAPH_IF_ELSE)
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


      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'correct_render', 'test_if_in_same_paragraph.docx'))
      our_render = Office::WordDocument.new(filename)

      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def test_if_else_in_tables
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'if_else_in_tables.docx'

      doc = Office::WordDocument.new(IF_ELSE_IN_TABLES)
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

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'correct_render', 'if_else_in_tables.docx'))
      our_render = Office::WordDocument.new(filename)
      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def test_nested_if_else_in_same_paragraph
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'nested_if_else_in_same_paragraph.docx'

      doc = Office::WordDocument.new(IN_SAME_PARAGRAPH_NESTED_IF_ELSE)
      template = Word::Template.new(doc)
      template.render(
        {'fields' =>
          {
            'a' => '2',
            'b' => 'haha',
            'c' => 'lol'
          }
        }, {do_not_render: true})
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'correct_render', 'nested_if_else_in_same_paragraph.docx'))
      our_render = Office::WordDocument.new(filename)
      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def test_if_with_and
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_if_with_and.docx'

      doc = Office::WordDocument.new(IF_WITH_AND)
      template = Word::Template.new(doc)
      template.render(
        {'fields' =>
          {
            'a' => '2',
            'b' => 'haha',
            'c' => 'lol'
          }
        }, {do_not_render: true})
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'correct_render', 'test_if_with_and.docx'))
      our_render = Office::WordDocument.new(filename)
      assert docs_are_equivalent?(correct, our_render)
    end
  end
end
