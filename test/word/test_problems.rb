require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'

class TestProblems < Test::Unit::TestCase
  TEST_ACTIVE_DIRECTORY_TEMPLATE = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'active_directory_woes_for_if_test.docx')
  TEST_SOLAR_PV_TEMPLATE = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'solar_pv.docx')

  #
  def test_active_directory_template
    file = File.new('active_directory_woes_for_if_test.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(TEST_ACTIVE_DIRECTORY_TEMPLATE)
    template = Word::Template.new(doc)
    template.render(
      {'fields' => {"Company__Contact___Contract_Info"=>
      [{"Company_Info"=>"",
        "Contract_Info"=>"",
        "Contact_Info"=>[{"Customer_Contact_Info"=>"", "Outgoing_Provider_Contact_Info"=>"", "TLG_Contact_Info"=>""}]}],
     "Active_Directory_Information"=>
      [{"Active_Directory_Domains"=>
         [{"Active_Directory_Domain_Name__FQDN_"=>"Pol",
           "Active_Directory_Domain_Name__NetBIOS_"=>"The",
           "Active_Directory_Domain_Administrator_Username"=>"Gigi",
           "Active_Directory_Domain_Administrator_Password"=>"ghh",
           "Directory_Services_Restore_Mode_Password"=>"ghh",
           "Service_Accounts"=>[{"Service_Account_Username"=>"Uhh", "Service_Account_Password"=>"yhh", "Service_Account_Notes_Description"=>"Guy"}]},
          {"Active_Directory_Domain_Name__FQDN_"=>"Yet"}]}],
     "Internet_Service_Providers"=>"",
     "Internet_Domains"=>"",
     "Web_Hosting"=>"",
     "E_Mail_Server_Provider_Information"=>"",
     "SSL_Certificates"=>"",
     "Remote_Access"=>"",
     "Backup___Disaster_Recovery"=>"",
     "Endpoint_Security"=>"",
     "Wireless"=>"",
     "Software___Licensing"=>"",
     "Hardware___Equipment"=>""}
        },{do_not_render: false})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'correct_render', 'active_directory_woes_for_if_test.docx'))
    our_render = Office::WordDocument.new(filename)

    assert docs_are_equivalent?(correct, our_render)
  ensure
    File.delete(filename)
  end

  def test_solar_pv_template
    file = File.new('solar_pv.docx', 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(TEST_SOLAR_PV_TEMPLATE)
    template = Word::Template.new(doc)
    template.render(
      {'fields' => {"Building_Name"=>"First Option", "Building_Location"=>"Somewhere", "Building_Address"=>"99 lolpie lane", "_1__Client_Details"=>[{"_1_1_1__Name___Facilities"=>"Aa", "_1_1_2__Position___Facilities"=>"Aa", "_1_1_3__Telephone___Facilities"=>"74", "_1_2_1__Name___Maintenance"=>"Aha", "_1_2_2__Position___Maintenance"=>"Aha", "_1_2_3__Telephone___Maintenance"=>"5454", "_1_3_1__Name___Building_Owner_Manager"=>"Bags", "_1_3_2__Position___Building_Owner_Manager"=>"Baja", "_1_3_3__Telephone___Building_Owner_Manager"=>"8454"}], "_2__Is_the_building_owned_by_City_of_Cape_Town_"=>"Yes", "_2_1__Comment"=>"Has", "_3__Parapit_Wall"=>[{"_3_1__Parapit_Wall_Height"=>"200 - 600 mm", "_3_2__Comments___Concrete_Roof"=>"Is", "_3_3__Please_take_a_photo_of_the_parapit_wall"=>"IMAGE"}], "_4__Roof_Segment"=>[{"Roof_Segment"=>"First Option", "Roof_Segment_Location"=>"UOu KNow", "_4_1_1__Access_to_roof"=>"Ladder", "_4_1_2__Description_of_access"=>"Is", "_4_1_3__Please_take_a_photo_of_the_access"=>"IMAGE", "_4_2_1_1__Dimensions"=>"Kiss", "_4_2_1_2__Sketch_roof_outline_and_dimension"=>"IMAGE", "_4_2_2__Orientation"=>"25", "_4_2_3__Roof_Tilt_Angle"=>"20 deg", "_4_2_4__Please_take_a_photo_of_the_roof"=>"IMAGE", "_4_3__Comments___Dimensions"=>"Ha", "_4_4_1__Lightning_Protection_Type"=>"Air Terminals", "_4_4_2__Please_take_a_photo_of_the_lightning_protection"=>"IMAGE", "_4_4_3__Air_Terminal_Height"=>"64", "_4_4_4__Please_take_a_photo_of_the_air_terminal"=>"IMAGE", "_4_4_5__Comments___Lightning_Protection"=>"Has", "_4_5_1__Shading_7_10am"=>"Severe", "_4_5_2__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_3__Shading_10am___2pm"=>"Moderate", "_4_5_4__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_5__Shading_2_6pm"=>"Moderate", "_4_5_6__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_7__Comments___Shading"=>"Uses"}, {"Roof_Segment"=>"First Option", "Roof_Segment_Location"=>"UOu KNow", "_4_1_1__Access_to_roof"=>"Scaffolding", "_4_1_2__Description_of_access"=>"It's", "_4_1_3__Please_take_a_photo_of_the_access"=>"IMAGE", "_4_2_1_1__Dimensions"=>"Aha", "_4_2_1_2__Sketch_roof_outline_and_dimension"=>"IMAGE", "_4_2_2__Orientation"=>"87", "_4_2_3__Roof_Tilt_Angle"=>"25 - 40 deg", "_4_2_4__Please_take_a_photo_of_the_roof"=>"IMAGE", "_4_3__Comments___Dimensions"=>"Aha", "_4_4_1__Lightning_Protection_Type"=>"Other -", "_4_4_2__Please_take_a_photo_of_the_lightning_protection"=>"IMAGE", "_4_4_5__Comments___Lightning_Protection"=>"Aha", "_4_5_1__Shading_7_10am"=>"Moderate", "_4_5_2__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_3__Shading_10am___2pm"=>"Moderate", "_4_5_4__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_5__Shading_2_6pm"=>"Moderate", "_4_5_6__Please_take_a_photo_to_show_shading"=>"IMAGE", "_4_5_7__Comments___Shading"=>"Shaun "}], "_5__Inverter"=>[{"_5_1__Inverter_Location"=>"Ha", "_5_2__Wall_space_for_inverters___Width"=>"54", "_5_3__Wall_space_for_inverters___Height"=>"644", "_5_4__Please_take_a_photo_of_inverter_location"=>"IMAGE", "_5_5__Comments___Inverters"=>"Hash"}, {"_5_1__Inverter_Location"=>"Ll", "_5_2__Wall_space_for_inverters___Width"=>"67", "_5_3__Wall_space_for_inverters___Height"=>"76", "_5_4__Please_take_a_photo_of_inverter_location"=>"IMAGE", "_5_5__Comments___Inverters"=>"Zhjs"}], "_6__Tie_in_DB"=>[{"_6_1_1__Tie_in_DB_Location"=>"Ha", "_6_1_2__DB_Name"=>"As", "_6_1_3__Main_or_Sub_DB"=>"Sub DB", "_6_1_4__Please_take_a_photo_of_DB_Location"=>"IMAGE", "_6_2_1__Main_Breaker_Size__A_"=>"77", "_6_2_2__Please_take_a_photo_of_the_main_breaker"=>"IMAGE", "_6_3_1__Fault_Level__kA_"=>"67", "_6_3_2__Please_take_a_photo_of_fault_level"=>"IMAGE", "_6_4_1__Spare_space_in_DB"=>"As", "_6_4_2__Please_take_photo_of_the_spare_DB_space"=>"IMAGE", "_6_5_1__Wall_space_for_new_DB___Width"=>"67", "_6_5_2__Wall_space_for_new_DB___Height"=>"94", "_6_5_3__Please_take_a_photo_of_the_available_wall_space"=>"IMAGE", "_6_6__Comments___Tie_in_DB"=>"Has"}, {"_6_1_1__Tie_in_DB_Location"=>"Aha", "_6_1_2__DB_Name"=>"Kayla", "_6_1_3__Main_or_Sub_DB"=>"Sub DB", "_6_1_4__Please_take_a_photo_of_DB_Location"=>"IMAGE", "_6_2_1__Main_Breaker_Size__A_"=>"976", "_6_2_2__Please_take_a_photo_of_the_main_breaker"=>"IMAGE", "_6_3_1__Fault_Level__kA_"=>"84543", "_6_3_2__Please_take_a_photo_of_fault_level"=>"IMAGE", "_6_4_1__Spare_space_in_DB"=>"Have ", "_6_4_2__Please_take_photo_of_the_spare_DB_space"=>"IMAGE", "_6_5_1__Wall_space_for_new_DB___Width"=>"76", "_6_5_2__Wall_space_for_new_DB___Height"=>"87", "_6_5_3__Please_take_a_photo_of_the_available_wall_space"=>"IMAGE", "_6_6__Comments___Tie_in_DB"=>"Has"}], "_7__Cable_Routing"=>[{"_7_1_1__Is_there_a_clear_path_between_roof_and_DB_room_"=>"true", "_7_1_2__Description_of_path"=>"Ajin", "_7_1_3__Please_take_a_photo_of_path"=>"IMAGE", "_7_2_1__Existing_Wireways"=>"Other -", "_7_2_2__Please_take_a_photo_of_the_existing_wireways"=>"IMAGE", "_7_3_1__Are_additional_wireways_required_"=>"true", "_7_3_2__Please_take_a_photo_where_wireways_are_required"=>"IMAGE", "_7_4__Comments___Cable_Routing"=>"Jana"}], "_8__Contractor_Site_Camp"=>[{"_8_1_1__Location"=>"lat=-33.863409, long=18.641795, alt=165.587158, hAccuracy=65.000000, vAccuracy=10.000000, timestamp=2016-10-04T09:44:57Z", "_8_1_2__Is_there_sufficient_space_for_2x_containers_"=>"true", "_8_1_3__Please_take_a_photo_of_site_camp_location"=>"IMAGE", "_8_2_1__Security"=>"Difficult to secure", "_8_2_2__Please_take_a_photo_of_security"=>"IMAGE", "_8_3__Comments___Contractor_Site_Camp"=>"An"}], "Detail_Sketches"=>[{"Sketch"=>"IMAGE", "Caption_Sketch"=>"He"}, {"Sketch"=>"IMAGE", "Caption_Sketch"=>"Hash"}], "General_Commentary"=>"Shhsxj", "Photo_of_Building_Side_1"=>"IMAGE", "Photo_of_Building_Side_2"=>"IMAGE", "Photo_of_Building_Side_3"=>"IMAGE", "Photo_of_Building_Side_4"=>"IMAGE"}},{do_not_render: false}
    )
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'correct_render', 'solar_pv.docx'))
    our_render = Office::WordDocument.new(filename)

    assert docs_are_equivalent?(correct, our_render)
  ensure
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
