require 'test/unit'
require 'date'
require 'office_docs'
require 'equivalent-xml'
require 'pry'

class TestProblems < Test::Unit::TestCase
  TEST_ACTIVE_DIRECTORY_TEMPLATE = File.join(File.dirname(__FILE__), '..', 'content', 'template', 'problems', 'active_directory_woes_for_if_test.docx')

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

  private

  def docs_are_equivalent?(doc1, doc2)
    xml_1 = doc1.main_doc.part.xml
    xml_2 = doc2.main_doc.part.xml
    EquivalentXml.equivalent?(xml_1, xml_2, { :element_order => true }) { |n1, n2, result| return false unless result }

    # TODO docs_are_equivalent? : check other doc properties

    true
  end
end
