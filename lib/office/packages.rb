require 'zip/zip'  # docs at http://rubyzip.sourceforge.net
require 'office/errors'
require 'office/parts'
require 'office/logger'

module Office
  class Package
    def initialize(filename)
      raise PackageError.new("cannot access '#{filename}' as a package file") unless File.file?(filename)

      @parts_by_name = {}
      Zip::ZipFile.open(filename) do |zip|
        zip.each do |entry| 
          part = Part.from_zip_entry(entry.name, entry.get_input_stream)
          @parts_by_name[part.name] = part
        end
      end
    end

    def debug_dump
      rows = @parts_by_name.values.collect { |p| ["#{p.class.name}", "#{p.name}", "#{p.content_type}"] }
      Logger.debug_dump_table("#{self.class.name} Parts", ["Class", "Name", "Content Type"], rows)
    end
  end

  class WordDocument < Package
    def initialize(filename)
      super(filename)
    end
  end
end
