require 'nokogiri' # docs at http://nokogiri.org
require 'RMagick'  # docs at http://studio.imagemagick.org/RMagick/doc
require 'office/constants'
require 'office/errors'
require 'office/logger'

module Office
  class Part
    attr_accessor :name
    attr_accessor :content_type

    def path_components
      name.split('/').values_at(1..-2)
    end

    def save(zip_output)
      zip_output.put_next_entry @name[1..-1] # strip off leading '/'
      zip_output << get_zip_content
    end

    def get_zip_content
      raise PackageError.new("incomplete implementation - get_zip_content for #{self.class.name}")
    end

    def get_comparison_content
      get_zip_content
    end
    
    def has_relationships?
      !@relationships.nil?
    end

    def set_relationships(relationships_part)
      raise "multiple relationship parts for '#{@name}'" unless @relationships.nil?
      @relationships = relationships_part
    end

    # returns RelationshipsPart
    def get_relationships
      @relationships
    end

    def get_relationship_by_id(id)
      @relationships.get_relationship_by_id(id)
    end

    def get_relationship_target(type)
      @relationships.get_relationship_target(type)
    end

    def add_relationship(part, type)
      @relationships.add(part, type)
    end

    def remove_relationships(part)
      return if self == part || @relationships.nil?
      @relationships.remove(part)
    end

    @@subclasses = []

    def self.from_entry(entry_name, entry_io, default_content_type = nil)
      part_name = (entry_name[0] == "/" ? "" : "/") + entry_name.downcase
      extension = part_name.split('.').last
      part_class = @@subclasses.find { |sc| sc.zip_extensions.include? extension } || UnknownPart

      begin
        part_class.parse(part_name, entry_io, default_content_type)
      rescue StandardError => e
        raise PackageError.new("failed to parse package #{part_class.name} '#{part_name}' - #{e}")
      end
    end

    def self.inherited(subclass)
      @@subclasses << subclass
    end

    def self.zip_extensions
      []
    end
  end

  class XmlPart < Part
    attr_accessor :xml # Nokogiri::XML::Document

    def initialize(part_name, xml_io, content_type)
      @name = part_name
      @xml = Nokogiri::XML::Document.parse(xml_io)
      @content_type = content_type
    end

    def get_zip_content
      @xml.to_xml(:indent => 0, :indent_text => nil).to_s
    end

    def self.parse(part_name, io, default_content_type)
      XmlPart.new(part_name, io, default_content_type || XML_CONTENT_TYPE)
    end

    def self.zip_extensions
      [ 'xml' ]
    end
  end

  class Relationship
    attr_accessor :id
    attr_accessor :type
    attr_accessor :target_name
    attr_accessor :target_part

    def initialize(id, type, target_name, target_part = nil)
      @id = id
      @type = type
      @target_name = target_name
      @target_part = target_part
    end

    def resolve_target_part(package, owner_name)
      full_name = @target_name[0] == "/" ? @target_name : owner_name[0, owner_name.rindex("/") + 1] + @target_name
      @target_part = package.get_part(full_name)
      Logger.warn "Failed to resolve relationship target '#{@target_name}' for '#{owner_name}'" if @target_part.nil?
    end

    def self.create_node(document, relationship_id, type, target)
      sheet_node = document.create_element("Relationship")
      sheet_node["Id"] = relationship_id
      sheet_node["Type"] = type
      sheet_node["Target"] = target
      sheet_node
    end
  end

  class RelationshipsPart < XmlPart
    def initialize(part_name, xml_io, content_type)
      super(part_name, xml_io, content_type)
      parse_relationships
    end

    def parse_relationships
      root = @xml.root
      raise PackageError.new("relationship part '#{@name}' has unexpected root node '#{root.name}") unless root.name == "Relationships"

      @relationships_by_id = {}
      root.children.each do |child|
        if "Relationship" == child.name
          id = child["Id"]
          raise PackageError.new("relationship part '#{@name}' has duplicate relationship ID '#{id}") if @relationships_by_id.has_key? id
          @relationships_by_id[id] = Relationship.new(id, child["Type"], child["Target"].downcase)
        else
          Logger.warn "Unrecognized element '#{child.name}' in relationships XML part" unless child.text? and child.blank?
        end
      end
    end

    def map_relationships(package)
      owner = resolve_relationships_owner(package)
      raise PackageError.new("relationship part '#{@name}' references a non-existent part") if owner.nil?

      @owner_name = owner == package ? "/" : owner.name
      @relationships_by_id.values.each { |r| r.resolve_target_part(package, @owner_name) }
      owner.set_relationships(self)
    end

    def resolve_relationships_owner(package)
      return package if @name == "/_rels/.rels"

      # names are of the form "/a/b/_rels/c.xml.rels"
      path_components = @name.split('/')
      valid_name = path_components.length > 1
      valid_name &&= path_components[path_components.length - 2] == "_rels"
      valid_name &&= path_components.last[-5, 5] == ".rels"
      raise PackageError.new("relationship part '#{@name}' name is invalid") unless valid_name

      path_components.delete_at(path_components.length - 2)
      path_components.last.chop!.chop!.chop!.chop!.chop! # Ruby is awesome!
      package.get_part(path_components.join("/"))
    end

    def get_relationship_by_id(id)
      @relationships_by_id[id]
    end
    
    def get_relationship_target(type)
      @relationships_by_id.values.each { |r| return r.target_part if r.type == type }
      nil
    end

    def add(part, type)
      relationship_id = next_free_relationship_id
      target = relative_path_from_owner(part.name)
      @xml.root.add_child(Relationship.create_node(@xml, relationship_id, type, target))
      @relationships_by_id[relationship_id] = Relationship.new(relationship_id, type, target, part)
      relationship_id
    end

    def remove(part)
      to_remove = []
      @relationships_by_id.values.each { |r| to_remove << r if r.target_part == part }
      to_remove.each do |r|
        @xml.root.at_xpath("/xmlns:Relationships/xmlns:Relationship[@Id='#{r.id}']").remove
        @relationships_by_id.delete(r.id)
      end
    end

    def next_free_relationship_id
      number = @relationships_by_id.size
      @relationships_by_id.keys.each do |k|
        matches = k.scan(/\ArId(\d+)\z/)
        number = [number, (matches.nil? || matches.empty? ? 0 : matches[0][0].to_i + 1)].max
      end
      "rId#{number}"
    end

    def relative_path_from_owner(part_name)
      owner_components = @owner_name.downcase.split('/')
      target_components = part_name.downcase.split('/')
      return part_name unless owner_components.first == target_components.first
      owner_components.each do |c|
        break unless target_components.first == c
        target_components.shift
      end
      target_components.join('/')
    end

    def debug_dump
      rows = @relationships_by_id.values.collect { |r| ["#{r.id}", "#{r.target_part.name}", "#{r.type}"] }
      title = "#{@owner_name == "/" ? "Package" : @owner_name} Relationships"
      Logger.debug_dump_table(title, ["ID", "Target", "Type"], rows)
    end

    def self.parse(part_name, io, default_content_type)
      RelationshipsPart.new(part_name, io, default_content_type || RELATIONSHIP_CONTENT_TYPE)
    end

    def self.zip_extensions
      [ 'rels' ]
    end
  end

  class ImagePart < Part
    attr_accessor :raw_blob
    attr_accessor :image # Magick::Image::Image

    def initialize(part_name, blob)
      @raw_blob = blob
      @name = part_name
      @image = Magick::Image::from_blob(blob).first
      @content_type = @image.mime_type
    end

    def get_comparison_content
      @image.signature
    end

    def get_zip_content
      @raw_blob
    end

    def self.parse(part_name, io, default_content_type)
      ImagePart.new(part_name, io.read)
    end

    def self.zip_extensions
      extensions = []
      Magick.formats do |format, attributes|
        # attributes indicate RMagick support for the format BRWA (native blob, read, write, multi-image)
        extensions << format.downcase if attributes.downcase[1, 2] = 'rw'
      end
      extensions
    end
  end

  class UnknownPart < Part
    attr_accessor :content

    def initialize(part_name, io, content_type)
      Logger.warn "Unknown Package Module: #{part_name}"
      @name = part_name
      @content = io.read
      @content_type = content_type
    end

    def get_zip_content
      @content
    end

    def self.parse(part_name, io, default_content_type)
      UnknownPart.new(part_name, io, default_content_type)
    end
  end
end
