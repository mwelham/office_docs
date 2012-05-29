require 'nokogiri' # docs at http://nokogiri.org
require 'RMagick'  # docs at http://studio.imagemagick.org/RMagick/doc
require 'office/errors'
require 'office/logger'

module Office
  class Part
    attr_accessor :name
    attr_accessor :content_type

    @@subclasses = []

    def self.from_zip_entry(entry_name, entry_io, default_content_type = nil)
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

    def self.parse(part_name, io, default_content_type)
      XmlPart.new(part_name, io, default_content_type || "application/xml")
    end

    def self.zip_extensions
      [ 'xml' ]
    end
  end

  class RelationshipsPart < XmlPart
    def self.parse(part_name, io, default_content_type)
      RelationshipsPart.new(part_name, io, default_content_type || "application/vnd.openxmlformats-package.relationships+xml")
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
      @content_type = @image.mime_type
    end

    def self.parse(part_name, io, default_content_type)
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

    def initialize(part_name, io, content_type)
      Logger.warn "Unknown Package Module: #{part_name}"
      @name = part_name
      @content = io.read
      @content_type = content_type
    end

    def self.parse(part_name, io, default_content_type)
      UnknownPart.new(part_name, io, default_content_type)
    end
  end
end
