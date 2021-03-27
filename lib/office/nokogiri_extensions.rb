require 'nokogiri'

module Nokogiri::XML::Searchable
  XMLNS = 'xmlns'.freeze
  XMLNS_COLON = "#{XMLNS}:".freeze
  TILDE = '~'.freeze
  COLON = ':'.freeze

  # convert xmlns= declarations to tag prefixes, eg
  # wpc:   from xmlns:wpc from xmlns:wpc"="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  # xmlns: from xmlns     from xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  # r:     from xmlns:r   from xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  private def namespace_colon namespace_decl
    namespace_decl.split(COLON).last << COLON
  end

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
  # ~ will be replaced with: whatever the default namespace is; or 'xmlns:'; or it
  # will do nothing if there are multiple namespaces, because you'll have to
  # differentiate anyway.
  #
  # TODO probably need some kind of fancy splatting and args extraction here to
  # fit in with the Searchable#xpath and Searchable#search
  def nspath xpath
    ns_xpath =
    case document.namespaces.size
    when 0
      # no namespace needed
      xpath.gsub TILDE, ''
    when 1
      # TODO probably need something like what Package.xpath_ns_prefix does:
      # def self.xpath_ns_prefix(node)
      #   node.nil? or node.namespace.nil? or node.namespace.prefix.blank? ? 'xmlns' : node.namespace.prefix
      # end

      # 'xmlns:wpc' from {"xmlns:wpc"="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"}
      ns, _url = namespaces.first
      xpath.gsub TILDE, namespace_colon(ns)
    else
      # use default xmlns: namespace if it exists, otherwise use whatever one is listed first
      if document.namespaces.key? XMLNS
        xpath.gsub TILDE, XMLNS_COLON
      else
        xpath.gsub TILDE, namespace_colon(namespaces.keys.first)
      end
    end

    # call in to the normal nokogiri method
    self.xpath ns_xpath
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
