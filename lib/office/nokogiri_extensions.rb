require 'nokogiri'

module Nokogiri::XML::Searchable
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
      xpath.gsub '~', 'xmlns:'
    when 1
      # TODO probably need something like what Package.xpath_ns_prefix does:
      # def self.xpath_ns_prefix(node)
      #   node.nil? or node.namespace.nil? or node.namespace.prefix.blank? ? 'xmlns' : node.namespace.prefix
      # end

      # 'wpc:' from {"xmlns:wpc"=>"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"}
      xmlns = namespaces.first.first

      # fortunately this also works for "xmlns" => "http://blablah"
      ns = xmlns.split(?:).last
      xpath.gsub '~', "#{ns}:"
    else
      # TODO maybe warning here
      xpath.gsub '~', ''
    end

    self.xpath ns_xpath
  end
end

# allow for >=ruby-2.7 pattern matching
class Nokogiri::XML::NodeSet
  def deconstruct; to_a end
end
