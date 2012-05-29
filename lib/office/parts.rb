require 'nokogiri' # docs at http://nokogiri.org
require 'RMagick'  # docs at http://studio.imagemagick.org/RMagick/doc
require 'office/errors'
require 'office/logger'

module Office
  class Part
    attr_accessor :name
    attr_accessor :content_type
    
    @@subclasses = []
    
    def self.from_zip_entry(entry_name, entry_io)
      part_name = entry_name.downcase
      extension = part_name.split('.').last
      part_class = @@subclasses.find { |sc| sc.zip_extensions.include? extension } || UnknownPart

      begin
        part_class.parse(part_name, entry_io)
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

    def initialize(part_name, xml_io)
      @name = part_name
      @content_type = "application/xml"
      @xml = Nokogiri::XML::Document.parse(xml_io)
    end
    
    def self.parse(part_name, io)
      XmlPart.new(part_name, io)
    end

    def self.zip_extensions
      [ 'xml' ]
    end
  end

  class RelationshipsPart < XmlPart
    def initialize(part_name, xml_io)
      super(part_name, xml_io)
      @content_type = "application/vnd.openxmlformats-package.relationships+xml"
    end
    
    def self.parse(part_name, io)
      RelationshipsPart.new(part_name, io)
    end

    def self.zip_extensions
      [ 'rels' ]
    end
  end
  
  class ImagePart < Part
    attr_accessor :image # Magick::Image::Image

    def initialize(part_name, image_list)
      @name = part_name
      @image = image_list.first
      @content_type = image.mime_type
    end
    
    def self.parse(part_name, io)
      ImagePart.new(part_name, Magick::Image::from_blob(io.read))
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
    
    def initialize(part_name, io)
      Logger.warn "Unknown Package Module: #{part_name}"
      @name = part_name
      @content = io.read
    end

    def self.parse(part_name, io)
      UnknownPart.new(part_name, io)
    end
  end
end
