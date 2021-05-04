require 'nokogiri'

module Nokogiri::XML::Searchable
  # DEPRECATED this should all be replaced in favour of nxpath, below. But no mandate for that right now.
  #
  # allow for ~ to be used in xpath expressions instead of xmlns:
  #
  # eg /~sst/~si/~t instead of /xmlns:sst/xmlns:si/xmlns:t
  #
  # Will also try to do the right thing if there are several namespace declarations.
  #
  # Rationale:
  #
  # libxml2 is pedantic about namespaces in xpath. Which is, quite frankly, a PITA. eg
  #
  #   doc.xpath '/xmlns:sst/xmlns:ssi/xmlns:t'
  #
  # So this entry point provides a way to say
  #
  #  doc.nspath '/~sst/~ssi/~t'
  #
  # Which is easier to type, easier to get right, and easier to read. And easy
  # to process because we don't have to do a full syntax parse to figure out which
  # identifiers are element names.
  #
  # probably need some kind of fancy splatting and args extraction here to
  # fit in with the Searchable#xpath and Searchable#search
  def nspath xpath
    self.xpath xpath.gsub /~(\w+)/, "node()[local-name() = '\\1']"
  end

  # implement xpath-2.0 //*:tag_name syntax meaning "any namespace with local-name() = 'tag_name'"
  def nxpath xpath
    self.xpath xpath.gsub /\*:(\w+)/, "node()[local-name() = '\\1']"
  end
end

# allow for >=ruby-2.7 pattern matching
class Nokogiri::XML::NodeSet
  def deconstruct; to_a end
end

class Nokogiri::XML::Document
  # convenience for create_element followed by builder
  def build_element name, **kwargs, &bld_blk
    create_element name, **kwargs do |node|
      Nokogiri::XML::Builder.with node, &bld_blk
    end
  end
end
