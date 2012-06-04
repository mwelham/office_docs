require 'zip/zip'  # docs at http://rubyzip.sourceforge.net
require 'office/errors'
require 'office/parts'
require 'office/logger'

module Office
  class Package
    def initialize(filename)
      raise PackageError.new("cannot access '#{filename}' as a package file") unless File.file?(filename)

      @filename = filename
      parse_parts
      map_relationships
    end

    def get_part(name)
      @parts_by_name[name]
    end

    def save(filename)
      if File.exists? filename
        backup_file = filename + ".bak"
        File.rename(filename, backup_file)
      end

      begin
        Zip::ZipOutputStream.open(filename) do |zip|
          @parts_by_name.values.each { |p| p.save(zip) }
        end
        File.delete(backup_file) unless backup_file.nil?
      rescue => e
        File.delete(filename) if File.exists? filename
        File.rename(backup_file, filename) unless backup_file.nil?
        raise e
      end
    end

    def parse_parts
      @parts_by_name = {}
      @default_content_types = {}
      @overriden_content_types = {}
      
      Zip::ZipFile.open(@filename) do |zip|
        entries = []
        zip.each do |e|
          if "/[Content_Types].xml".casecmp(e.name[0] == "/" ? e.name : "/" + e.name) == 0
            parse_content_types(e)
          else
            entries << e
          end
        end
        raise PackageError.new("package '#{@filename}' is missing content types part") if @parts_by_name.empty?
        entries.each { |e| parse_zip_entry(e) }
      end
    end

    def parse_zip_entry(zip_entry)
      extension = zip_entry.name.split('.').last.downcase
      part = Part.from_zip_entry(zip_entry.name, zip_entry.get_input_stream, @default_content_types[extension])
      part.content_type = @overriden_content_types[part.name] || part.content_type
      @parts_by_name[part.name] = part
      part
    end

    def parse_content_types(zip_entry)
      part = parse_zip_entry(zip_entry)
      type_node = part.xml.root
      raise PackageError.new("package '#{@filename}' has unexpected root node '#{type_node.name}") unless type_node.name == "Types"
      
      @default_content_types = {}
      type_node.children.each do |child|
        case child.name
        when "Default"
          @default_content_types[child["Extension"].downcase] = child["ContentType"]
        when "Override"
          @overriden_content_types[child["PartName"].downcase] = child["ContentType"]
        else
          Logger.warn "Unrecognized element '#{child.name}' in content types XML part"
        end
      end
    end

    def map_relationships
      @parts_by_name.values.each { |p| p.map_relationships(self) if p.instance_of? RelationshipsPart }
      Logger.warn "package '#{@filename}' is missing package-level relationships" if @relationships.nil?
    end

    def set_relationships(relationships_part)
      raise "multiple package-level relationship parts for package '#{@filename}'" unless @relationships.nil?
      @relationships = relationships_part
    end

    def get_relationship_target(type)
      raise "package '#{@filename}' is missing package-level relationships" if @relationships.nil?
      @relationships.get_relationship_target(type)
    end

    def debug_dump
      rows = @parts_by_name.values.collect { |p| ["#{p.class.name}", "#{p.name}", "#{p.content_type}"] }
      Logger.debug_dump_table("#{self.class.name} Parts", ["Class", "Name", "Content Type"], rows)

      @relationships.debug_dump unless @relationships.nil?
      @parts_by_name.values.each { |p| p.get_relationships.debug_dump if p.has_relationships? }
    end
  end
end
