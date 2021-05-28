module FixtureFiles
  def self.content_path
    Pathname(__dir__) + "../test/content"
  end

  def self.fixtures_path
    Pathname(__dir__) + "fixtures"
  end

  module Book; end
  module Doc; end
  module Image; end
  module Xml; end
  module Yaml; end
  module Json; end
  module Txt; end

  (content_path.glob('*.*') + fixtures_path.glob('*.*')).each do |path|
    the_mod =
    case path.extname
    when '.xlsx'; Book
    when '.docx'; Doc
    when '.jpg'; Image
    when '.xml'; Xml
    when '.yml'; Yaml
    when '.json'; Json
    when '.txt'; Txt
    else; raise "Unknown module for #{path}"
    end

    const_name = path.basename(path.extname).to_s.gsub(/[[:punct:]]/, ?_).upcase
    the_mod.const_set const_name, path.realpath.to_s
  end
end

module Reload
  def reload document, filename = nil, &blk
    Dir.mktmpdir do |dir|
      filename = File.join dir, (filename || File.basename(document.filename))
      document.save filename
      yield document.class.new(filename)
    end
  end

  alias reload_workbook reload
  alias reload_document reload
end

ReloadWorkbook = Reload
