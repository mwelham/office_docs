module LoopOverParagraphsTest
  IN_DIFFERENT_PARAGRAPH_FOR_LOOP = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'in_different_paragraph_for_loop_test.docx')
  IN_DIFFERENT_PARAGRAPH_FOR_LOOP_COMPLEX = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'in_different_paragraph_complex_loop_test.docx')
  LOOP_IN_LOOP_IN_DIFFERENT_PARAGRAPH = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'test_loop_in_loop_in_different_paragraph.docx')
  LOOP_OVER_GRAPHIC = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'loop_over_graphic_test.docx')
  LOOP_OVER_LIST = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'loop_over_list_test.docx')

  COMPLEX_FOR_LOOP = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'complex_for_loop_word_template.docx')

  #DIFFERENT PARAGRAPH LOOPS
  #
  #
  #
  def test_loop_in_different_paragraph_with_blank_group
    file = File.new('test_loop_in_different_paragraph_with_blank_group.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(IN_DIFFERENT_PARAGRAPH_FOR_LOOP)
    template = Word::Template.new(doc)
    template.render(
      {'fields' =>
        {'Group' => []
      }
    }, {do_not_render: true})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'in_different_paragraph_blank_group.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end

  def test_loop_in_different_paragraph
    file = File.new('test_loop_in_different_paragraph.docx', 'w')
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
    }, {do_not_render: true})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'in_different_paragraphs.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end

  def test_complex_loop_in_different_paragraph
    file = File.new('test_complex_loop_in_different_paragraph.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(IN_DIFFERENT_PARAGRAPH_FOR_LOOP_COMPLEX)
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

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'test_complex_loop_in_different_paragraph.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end

  def test_loop_in_loop
    file = File.new('test_loop_in_loop_in_different_paragraph.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(LOOP_IN_LOOP_IN_DIFFERENT_PARAGRAPH)
    template = Word::Template.new(doc)
    template.render(
      {'fields' =>
        {'Group' => [
          {'Boss' => 'Gandalf', 'Employees' => [{'Name' => 'Frodo'}, {'Name' => 'Sam'}]},
          {'Boss' => 'Leya', 'Employees' => [{'Name' => 'Luke'}, {'Name' => 'Han'}, {'Name' => 'Chewbaka'}]},
          {'Boss' => 'Morpheus', 'Employees' => [{'Name' => 'Neo'}, {'Name' => 'Trinity'}]},
          {'Boss' => 'The Dude', 'Employees' => []}
        ]
      }
    }, {do_not_render: true})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'test_loop_in_loop_in_different_paragraph.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end

  def test_loop_in_loop_with_render
    file = File.new('test_loop_in_loop_in_different_paragraph_with_render.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(LOOP_IN_LOOP_IN_DIFFERENT_PARAGRAPH)
    template = Word::Template.new(doc)
    template.render(
      {'fields' =>
        {'Group' => [
          {'Boss' => 'Gandalf', 'Employees' => [{'Name' => 'Frodo'}, {'Name' => 'Sam'}]},
          {'Boss' => 'Leya', 'Employees' => [{'Name' => 'Luke'}, {'Name' => 'Han'}, {'Name' => 'Chewbaka'}]},
          {'Boss' => 'Morpheus', 'Employees' => [{'Name' => 'Neo'}, {'Name' => 'Trinity'}]},
          {'Boss' => 'The Dude', 'Employees' => []}
        ]
      }
    })
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'test_loop_in_loop_in_different_paragraph_with_render.docx'))
    our_render = Office::WordDocument.new(filename)
    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end

  def test_loop_over_graphic
    file = File.new('loop_over_graphic_test.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(LOOP_OVER_GRAPHIC)
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

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'loop_over_graphic_test.docx'))
    our_render = Office::WordDocument.new(filename)

    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end

  def test_loop_over_graphic
    file = File.new('loop_over_list_test.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(LOOP_OVER_LIST)
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

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'loop_over_list_test.docx'))
    our_render = Office::WordDocument.new(filename)

    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end

  def test_complex_loop
    file = File.new('complex_for_loop_word_template.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(COMPLEX_FOR_LOOP)
    template = Word::Template.new(doc)
    template.render(
      {'fields' => {"Inspector_Name"=>"Dr Matt",
       "Sites"=>
        [{"Name"=>"Site A",
          "Number_of_cool_things_at_the_site"=>"42",
          "Checks"=>
           [{"Check"=>"First Option", "Tools_found"=>[{"Tool_Name"=>"Bob"}, {"Tool_Name"=>"Hammer"}]},
            {"Check"=>"2", "Tools_found"=>[{"Tool_Name"=>"Shovel"}]},
            {"Check"=>"3"},
            {"Check"=>"4", "Tools_found"=>[{"Tool_Name"=>"A"}, {"Tool_Name"=>"B"}, {"Tool_Name"=>"C"}, {"Tool_Name"=>"D"}]}],
          "Employees"=>
           [{"Name"=>"Pop", "Rating"=>"Good"},
            {"Name"=>"Cap", "Rating"=>"Average"}]},
         {"Name"=>"Site B",
          "Number_of_cool_things_at_the_site"=>"2",
          "Checks"=>[{"Check"=>"6", "Tools_found"=>[{"Tool_Name"=>"Pew"}, {"Tool_Name"=>"Shoe"}]}]}]
          }
        },{do_not_render: true})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'complex_for_loop_word_template.docx'))
    our_render = Office::WordDocument.new(filename)

    assert docs_are_equivalent?(correct, our_render)

    File.delete(filename)
  end
end
