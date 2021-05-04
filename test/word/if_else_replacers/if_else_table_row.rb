module IfElseTableRowTest
  IN_TABLE_ROW_DIFFERENT_CELL = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'if_else_same_row.docx')

  #TABLE ROW LOOPS
  #
  #
  #
  def test_loop_in_table_same_row_different_cells
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'if_else_same_row.docx'

      doc = Office::WordDocument.new(IN_TABLE_ROW_DIFFERENT_CELL)
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

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'if_else', 'correct_render', 'if_else_same_row.docx'))
      our_render = Office::WordDocument.new(filename)
      assert docs_are_equivalent?(correct, our_render)
    end
  end

end
