module BookFiles
  def self.content_path
    Pathname(__dir__) + "../test/content"
  end

  content_path.children.each do |path|
    next unless path.to_s.end_with? '.xlsx'
    const_name = path.basename('.xlsx').to_s.gsub(/[[:punct:]]/, ?_).upcase
    const_set const_name, path.realpath.to_s
  end
end

# TODO generalise
module DocFiles
  def self.content_path
    Pathname(__dir__) + "../test/content"
  end

  content_path.children.each do |path|
    next unless path.to_s.end_with? '.docx'
    const_name = path.basename('.docx').to_s.gsub(/[[:punct:]]/, ?_).upcase
    const_set const_name, path.realpath.to_s
  end
end

# TODO generalise
module ImageFiles
  def self.content_path
    Pathname(__dir__) + "../test/content"
  end

  content_path.children.each do |path|
    next unless path.to_s.end_with? '.jpg'
    const_name = path.basename('.jpg').to_s.gsub(/[[:punct:]]/, ?_).upcase
    const_set const_name, path.realpath.to_s
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
