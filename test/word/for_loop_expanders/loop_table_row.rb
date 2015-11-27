module LoopTableRowTest
  IN_TABLE_ROW_DIFFERENT_CELL = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'in_table_same_row_test.docx')

  #TABLE ROW LOOPS
  #
  #
  #
  def test_loop_in_table_same_row_different_cells
    file = File.new('test_simple_rows_loop_doc.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(IN_TABLE_ROW_DIFFERENT_CELL)
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

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'test_simple_rows_loop_doc.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end
end
