module PackageDebug
  refine Office::Package do
    def parts; @parts_by_name end

    def each_zip_entry &blk
      # fail because of refinement ...
      # return enum_for __method__ unless block_given?

      # ... so do it a bit clunkier
      Zip::File.open(@filename) do |zip|
        if block_given?
          zip.each &blk
        else
          zip.enum_for :each
        end
      end
    end

    def incl_tree_name file_name
      file_name =~ /(xml|rels)$/ && file_name !~ /styles|settings|theme|font/
    end

    # Create a giant xml for entire zip file.
    # Convert zip file structure to xml, and parse contents into child nodes
    # default parent to an empty but rooted doc
    def ztree parent = Nokogiri::XML.parse('<ztree/>').root
      Zip::File.open(@filename) do |zip|
        zip.each do |ze|
          node = parent.document.create_element ze.ftype.to_s
          node[:name] = file_name = ze.name.to_s.downcase
          # node[:size] = ze.size.to_s
          # these are always 1980-01-01 00:00:00 so pretty useless
          # node[:time] = ze.time.to_s
          # node[:mtime] = ze.mtime.to_s

          # exclude/include useful stuff. Should be a keyword or lambda.
          if incl_tree_name(file_name)
            contents = Nokogiri::XML.parse(ze.get_input_stream)&.root || parent.document.create_element('empty')
            # this (sometimes?) ensures that indentation is correct, node << contents does not
            contents.parent = node
          end

          node.parent = parent
        end
      end
      parent
    end

    # The parts sorted in current zip order.
    # raise unless it has a file
    def zorted_parts
      # establish the sort order
      zort = each_zip_entry.map{|zip_entry| "/#{zip_entry.name}".downcase}.each_with_index.to_h
      # sort parts same as zip entries. unknowns at the end in arbitrary order
      parts.sort{|(k1,_v1), (k2,_v2)| (zort[k1.downcase] || Float::INFINITY) <=> (zort[k2.downcase] || Float::INFINITY) }.to_h
    end

    # Create a giant xml for entire part list.
    # Basically a mirror to what self#save does but to an xml instead of a zip
    def xtree parent = Nokogiri::XML.parse('<xtree/>').root
      zorted_parts.each do |name, part|
        node = parent.document.create_element 'file'
        node[:name] = name.gsub(%r|^/|, '').downcase
        if incl_tree_name(name) && part.respond_to?(:xml)
          contents = part.xml.root.clone
          contents.parent = node
        end
        node.parent = parent
      end
      parent
    end
  end
end

