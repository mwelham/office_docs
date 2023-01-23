module TemplateTestHelper
  def check_template(path_to_template, path_to_correct_render, options = {})
    Dir.mktmpdir do |dir|
      filename = File.join dir, path_to_template.split('/').last

      doc = Office::WordDocument.new(path_to_template)
      template = Word::Template.new(doc)
      render_params = options[:render_params] || {}
      do_not_render = options[:do_not_render].nil? ? false : options[:do_not_render]
      template.render(render_params, {do_not_render: do_not_render})
      template.word_document.save(filename)

      assert File.file?(filename)
      assert File.stat(filename).size > 0

      binding.pry if options[:pry]

      correct = Office::WordDocument.new(path_to_correct_render)
      our_render = Office::WordDocument.new(filename)

      assert docs_are_equivalent?(correct, our_render)
    end
  end

  def image_to_test_with
    Magick::ImageList.new File.join(File.dirname(__FILE__), '..', 'content', 'test_image.jpg')
  end

  def docs_are_equivalent?(doc1, doc2)
    xml_1 = doc1.main_doc.part.xml
    xml_2 = doc2.main_doc.part.xml
    EquivalentXml.equivalent?(xml_1, xml_2, { :element_order => true }) { |n1, n2, result| return false unless result }

    # TODO docs_are_equivalent? : check other doc properties

    true
  end
end
