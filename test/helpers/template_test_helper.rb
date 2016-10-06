module TemplateTestHelper
  def check_template(path_to_template, path_to_correct_render, options = {})
    template_name = path_to_template.split('/').last
    file = File.new(template_name, 'w')
    file.close
    filename = file.path

    doc = Office::WordDocument.new(path_to_template)
    template = Word::Template.new(doc)
    render_params = options[:render_params] || {}
    do_not_render = options[:do_not_render].nil? ? false : options[:do_not_render]
    template.render(render_params, {do_not_render: do_not_render})
    template.word_document.save(filename)

    assert File.file?(filename)
    assert File.stat(filename).size > 0

    correct = Office::WordDocument.new(path_to_correct_render)
    our_render = Office::WordDocument.new(filename)

    assert docs_are_equivalent?(correct, our_render)
  ensure
    File.delete(filename)
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
