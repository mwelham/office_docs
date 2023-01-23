module LoopOverParagraphsTest
  require 'helpers/template_test_helper'

  IN_DIFFERENT_PARAGRAPH_FOR_LOOP = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'in_different_paragraph_for_loop_test.docx')
  IN_DIFFERENT_PARAGRAPH_FOR_LOOP_COMPLEX = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'in_different_paragraph_complex_loop_test.docx')
  LOOP_IN_LOOP_IN_DIFFERENT_PARAGRAPH = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'test_loop_in_loop_in_different_paragraph.docx')
  LOOP_OVER_GRAPHIC = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'loop_over_graphic_test.docx')
  LOOP_OVER_LIST = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'loop_over_list_test.docx')
  LOOP_OVER_IMAGE = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'loop_over_image.docx')

  COMPLEX_FOR_LOOP = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'complex_for_loop_word_template.docx')

  #DIFFERENT PARAGRAPH LOOPS
  #
  #
  #
  def test_loop_in_different_paragraph_with_blank_group
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_loop_in_different_paragraph_with_blank_group.docx'

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
    end
  end

  def test_loop_in_different_paragraph
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_loop_in_different_paragraph.docx'

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
    end
  end

  def test_complex_loop_in_different_paragraph
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_complex_loop_in_different_paragraph.docx'

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
    end
  end

  def test_loop_in_loop
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_loop_in_loop_in_different_paragraph.docx'

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
    end
  end

  def test_loop_in_loop_with_render
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_loop_in_loop_in_different_paragraph_with_render.docx'

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
    end
  end

  def test_loop_over_graphic
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'loop_over_graphic_test.docx'

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
    end
  end

  def test_loop_over_list
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'loop_over_list_test.docx'

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
    end
  end

  def test_loop_over_image
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'loop_over_image.docx'

      doc = Office::WordDocument.new(LOOP_OVER_IMAGE)
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

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'loop_over_image.docx'))
      our_render = Office::WordDocument.new(filename)

      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def test_complex_loop
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'complex_for_loop_word_template.docx'

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
    end
  end

  def test_complex_loop_with_render
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'complex_for_loop_word_template_with_render.docx'

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
          },{do_not_render: false})
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'complex_for_loop_word_template_with_render.docx'))
      our_render = Office::WordDocument.new(filename)

      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def test_blank_loop
    path_to_template = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'test_blank_loop_in_different_paragraph.docx')
    path_to_correct_render = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'test_blank_loop_in_different_paragraph.docx')
    render_params = {
      'fields' => {
      }
    }
    check_template(path_to_template, path_to_correct_render, {render_params: render_params})
  end
end
