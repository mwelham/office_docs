module LoopTableRowTest
  IN_TABLE_ROW_DIFFERENT_CELL = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'in_table_same_row_test.docx')
  PROBLEM_TURN_IN_ORDER = File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'problems', 'Turn In Order.docx')

  #TABLE ROW LOOPS
  #
  #
  #
  def test_loop_in_table_same_row_different_cells
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_simple_rows_loop_doc.docx'

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
    end
  end

  def test_loop_in_table_turn_in_order
    Dir.mktmpdir do |dir|
      filename = File.join dir, 'test_problem - Turn in Order.docx'

      doc = Office::WordDocument.new(PROBLEM_TURN_IN_ORDER)
      template = Word::Template.new(doc)
      template.render(turn_in_order_fields)
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0

      correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', '..', 'content', 'template', 'for_loops', 'correct_render', 'test_problem - Turn in Order.docx'))
      our_render = Office::WordDocument.new(filename)
      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def turn_in_order_fields
    {'fields' =>
        {"Date___Time"=>"2016-03-01 15:23:15",
         "TRU_SalesRep"=>"Andrea Cooper",
         "Andrea_Cooper_Account_List"=>"First Option",
         "Manager_Name"=>"I",
         "Store_Visit_Type"=>"Store Information",
         "Store_Information"=>
          [{"Trumps_SI_ProdDist"=>"My Organics",
            "PD_MO"=>[{"Sales_Rep"=>"greg@iglink.com.au", "PD_MO_001"=>"yes", "PD_MO_002"=>"no", "PD_MO_003"=>"yes", "PD_MO_004"=>"yes"}],
            "Shelf_Management"=>
             [{"Shelf_Mana_Group"=>
                [{"Department"=>"Produce", "Main_Product_Location"=>"Hangsell Wing", "ShelfMana_Photo_"=>"IMAGE"},
                 {"Department"=>"Grocery", "Main_Product_Location"=>"6 Foot Section", "ShelfMana_Photo_"=>"IMAGE"}]}],
            "Point_Of_Sale__POS_"=>[{"POS_Location"=>"Chrome Stand", "Image_1_Photo_"=>"IMAGE", "Image_2_Photo_"=>"IMAGE"}],
            "Opposition_Check"=>
             [{"Opposition_Ranging_Detail"=>"B",
               "Opposition_New_Lines_Detail"=>"P",
               "Opposition_New_Lines_Photo_"=>"IMAGE",
               "Opposition_Displays_Detail"=>"Zz"}],
            "Sales_Rep_Email"=>"greg@iglink.com.au",
            "Forward_a_copy_of_Shelf_Management_report_"=>"DBarwick@trumps.com.au"}],
         "ManaSign_photo_"=> test_image,
         "RepSign_photo_"=> test_image,
         "Time_Out"=>"15:25:02"}
    }
  end

  def test_image
    Magick::ImageList.new File.join(File.dirname(__FILE__), '..', '..', 'content', 'test_image.jpg')
  end
end
