module BookFiles
  (Pathname(__dir__) + "../test/content").children.each do |path|
    next unless path.to_s.end_with? '.xlsx'
    const_name = path.basename('.xlsx').to_s.gsub(/[[:punct:]]/, ?_).upcase
    const_set const_name, path.realpath.to_s
  end
end

module ReloadWorkbook
  def reload_workbook workbook, filename = nil, &blk
    Dir.mktmpdir do |dir|
      filename = File.join dir, (filename || File.basename(workbook.filename))
      workbook.save filename
      yield Office::ExcelWorkbook.new(filename)
    end
  end
end
