require 'zip'  # docs at http://rubyzip.sourceforge.net
require 'office/errors'
require 'office/parts'
require 'office/logger'

module Office
  class Package
    attr_accessor :filename

    def initialize(filename)
      raise PackageError.new("cannot access '#{filename}' as a package file") unless File.file?(filename)

      @filename = filename
      parse_parts
      map_relationships
    end

    def to_data
      Dir.mktmpdir do |dir|
        # generate a unique-enough filename
        tmp_filename = @filename || (now = Time.now; "#{now.to_i}.#{now.tv_nsec}")
        path = File.join dir, File.basename(tmp_filename)
        save path
        File.open(path, 'rb:ASCII-8BIT', &:read)
      end
    end

    def get_part(name)
      @parts_by_name[name]
    end

    def get_part_names
      @parts_by_name.keys
    end

    def add_part(name, content_io, content_type)
      PackageError.new("part name cannot be empty") if name.nil? or name.empty?
      PackageError.new("package already contains a part with name '#{name}'") if @parts_by_name.has_key? name
      add_content_type_override(name, content_type)
      part = Part.from_entry(name, content_io)
      @parts_by_name[name] = part
      part
    end

    def remove_part(part)
      return if part.nil?
      raise PackageError.new("part not found in package") unless @parts_by_name[part.name] == part

      @parts_by_name.values.each { |p| p.remove_relationships(part) }
      remove_content_type_override(part.name)
      @parts_by_name.delete(part.name)
    end

    def unused_part_identifier(prefix)
      i = 1
      until (@parts_by_name.keys.index { |name| name.start_with? "#{prefix}#{i}" }).nil?
        i = i + 1
      end
      i
    end

    def save(filename)
      if File.exist? filename
        backup_file = filename + ".bak"
        File.rename(filename, backup_file)
      end

      begin
        Zip::OutputStream.open(filename) do |zip|
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

      Zip::File.open(@filename) do |zip|
        entries = []
        zip.each do |e|
          if "/[Content_Types].xml".casecmp(e.name[0] == "/" ? e.name : "/" + e.name) == 0
            parse_content_types(e)
          else
            entries << e
          end
        end
        raise PackageError.new("package '#{@filename}' is missing content types part") if @parts_by_name.empty?
        entries.each do |e|
          parse_zip_entry(e) unless e.directory?
        end
      end
    end

    def parse_zip_entry(zip_entry)
      extension = zip_entry.name.split('.').last.downcase
      part = Part.from_entry(zip_entry.name, zip_entry.get_input_stream, @default_content_types[extension])
      part.content_type = @overriden_content_types[part.name] || part.content_type
      @parts_by_name[part.name] = part
      part
    end

    def add_content_type_override(name, content_type)
      return if content_type.nil? or content_type.empty?
      return if @default_content_types[name.split('.').last.downcase] == content_type

      node = @content_types_part.xml.create_element('Override')
      node["PartName"] = name
      node["ContentType"] = content_type
      @content_types_part.xml.root.add_child(node)
      @overriden_content_types[name.downcase] = content_type
    end

    def remove_content_type_override(name)
      ns_prefix = Package.xpath_ns_prefix(@content_types_part.xml.root)
      nodes = @content_types_part.xml.root.xpath("/#{ns_prefix}:Types/#{ns_prefix}:Override[@PartName='#{name}']")
      nodes.each { |n| n.remove } unless nodes.nil?
      @overriden_content_types.delete(name.downcase)
    end

    def parse_content_types(zip_entry)
      @content_types_part = parse_zip_entry(zip_entry)
      type_node = @content_types_part.xml.root
      raise PackageError.new("package '#{@filename}' has unexpected root node '#{type_node.name}") unless type_node.name == "Types"

      @default_content_types = {}
      type_node.children.each do |child|
        case child.name
        when "Default"
          @default_content_types[child["Extension"].downcase] = child["ContentType"]
        when "Override"
          @overriden_content_types[child["PartName"].downcase] = child["ContentType"]
        else
          Logger.warn "Unrecognized element '#{child.name}' in content types XML part" unless child.text? and child.blank?
        end
      end
    end

    def map_relationships
      @parts_by_name.values.each { |p| p.map_relationships(self) if p.instance_of? RelationshipsPart }
      Logger.warn "package '#{@filename}' is missing package-level relationships" if @relationships.nil?
    end

    def set_relationships(relationships_part)
      raise "multiple package-level relationship parts for package '#{@filename}'" if instance_variable_defined?(:@relationships)
      @relationships = relationships_part
    end

    def get_relationship_targets(type)
      raise "package '#{@filename}' is missing package-level relationships" if @relationships.nil?
      @relationships.get_relationship_targets(type)
    end

    # make sure that a maybe-new part has a related Relationships entry in the relevant rels file
    def ensure_relationships part
      unless part.has_relationships?
        content = StringIO.new %|<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>|
        rels_part = add_part(part.rels_name, content, RELATIONSHIP_CONTENT_TYPE)
        rels_part.map_relationships(self)
      end
    end

    # return rel_id for new relationship
    # MAYBE can lookup rel_type from somewhere? Will always be related to dst_part anyway...?
    def add_relationship(src_part, dst_part, rel_type)
      ensure_relationships src_part
      src_part.add_relationship(dst_part, rel_type)
    end

    # part is the Office::Part to which the image should be added
    # image will be added as a part
    # rel will be added from part to the new image_part
    def add_image_part_rel(image, part)
      image_part = add_image_part(image, part.path_components)
      relationship_id = add_relationship(part, image_part, IMAGE_RELATIONSHIP_TYPE)

      [relationship_id, image_part]
    end

    # image will be added as a part underneath /<path_components>/media/imageX.<imgext>
    # returns an ImagePart instance
    def add_image_part(image, path_components)
      prefix = File.join ?/, path_components, 'media/image'

      # unused_part_identifier is 1..n : Integer
      # .extension comes from the image
      part_name = "#{prefix}#{unused_part_identifier prefix}.#{image.format.downcase}"

      add_part(part_name, StringIO.new(image.to_blob), image.mime_type)
    end

    def debug_dump
      rows = @parts_by_name.values.collect { |p| ["#{p.class.name}", "#{p.name}", "#{p.content_type}"] }
      Logger.debug_dump_table("#{self.class.name} Parts", ["Class", "Name", "Content Type"], rows)

      @relationships.debug_dump unless @relationships.nil?
      @parts_by_name.values.each { |p| p.get_relationships.debug_dump if p.has_relationships? }
    end

    # don't depend on activesupport
    def self.blank?( obj )
      obj.nil? || obj == ''
    end

    def self.xpath_ns_prefix(node)
      node.nil? or node.namespace.nil? or blank?(node.namespace.prefix) ? 'xmlns' : node.namespace.prefix
    end
  end

  class PackageComparer
    def self.are_equal?(path_1, path_2)
      package_1 = Package.new(path_1)
      package_2 = Package.new(path_2)

      part_names_1 = package_1.get_part_names
      part_names_2 = package_2.get_part_names

      return false unless (part_names_1 - part_names_2).empty?
      return false unless (part_names_2 - part_names_1).empty?

      part_names_1.each do |name|
        return false unless are_parts_equal?(package_1.get_part(name), package_2.get_part(name))
      end
      true
    end

    def self.are_parts_equal?(part_1, part_2)
      return false unless part_1.class == part_2.class
      return false unless part_1.name == part_2.name
      return false unless part_1.content_type == part_2.content_type

      content_1 = part_1.get_comparison_content.to_s
      content_2 = part_2.get_comparison_content.to_s
      content_1 == content_2
    end
  end
end
