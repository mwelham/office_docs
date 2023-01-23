require_relative 'spec_helper'

require_relative '../lib/office/package'
require_relative '../lib/office/word'
require_relative '../lib/office/excel'

require_relative 'modules.rb'
require_relative 'package_debug.rb'

describe Office::Package do
  describe '#add_image_part_rel' do
    include Reload
    using PackageDebug

    let :image do Magick::ImageList.new FixtureFiles::Image::TEST_IMAGE end
    let :word do Office::WordDocument.new FixtureFiles::Doc::IMAGE_REPLACEMENT_TEST end
    let :book do Office::ExcelWorkbook.new FixtureFiles::Book::IMAGE_TEST end
    let :simple do Office::ExcelWorkbook.new FixtureFiles::Book::SIMPLE_TEST end

    def main_part document
      candidates =
      case document
      when Office::WordDocument; document.parts.filter{|name, _part| name =~ /word.document/}
      when Office::ExcelWorkbook; document.parts.filter{|name, _part| name =~ /xl.workbook/}
      end

      raise "must be one" unless candidates.size == 1
      # document.parts is a hash, so fetch the value from the first pair
      _name, part = candidates.first
      part
    end

    docs = {
      word: Office::WordDocument.new(FixtureFiles::Doc::IMAGE_REPLACEMENT_TEST),
      book: Office::ExcelWorkbook.new(FixtureFiles::Book::IMAGE_TEST),
      simple: Office::ExcelWorkbook.new(FixtureFiles::Book::SIMPLE_TEST),
      empty: Office::WordDocument.new(FixtureFiles::Doc::EMPTY),
    }

    docs.each do |name, doc|
      describe name do
        describe '#ensure_relationship' do
          let :part do main_part doc end
          let :image_part do doc.add_image_part image, part.path_components end

          it 'adds rels/.rels or whatever its called' do
            # preconditions
            # no rels file exists
            doc.parts[image_part.rels_name].should be_nil

            doc.ensure_relationships image_part

            # rels file now exists
            doc.parts[image_part.rels_name].should be_a(Office::RelationshipsPart)

            # contains a Relationships tag but nothing else
            rel = doc.parts[image_part.rels_name].xml
            rel.nxpath('/*:Relationships').size.should == 1
            rel.nxpath('/*:Relationships/node()').size.should == 0
          end
        end

        # eg Types/Override PartName="/xl/media/image1.jpeg" ContentType="image/jpeg"
        # or Types/Default Extension="jpeg" ContentType="image/jpeg"
        it 'jpeg has content type' do
          part = main_part doc
          rel_id, image_part = doc.add_image_part_rel image, part

          # overrides not necessary for jpeg because they already exist
          # node_set = content_types.xml.nxpath(%|/*:Types/*:Override/@PartName[text() = '#{image}']|)

          content_types = doc.parts["/[content_types].xml"]
          node_set = content_types.xml.nxpath(%|/*:Types/*:Default[@Extension = 'jpeg'][@ContentType = 'image/jpeg']|)
          node_set.size.should == 1
        end

        it 'adds rel of type image' do
          # here it doesn't really matter what part the image is added to, as long as it shows up in the saved zip/xml
          rel_id, _image_part = doc.add_image_part_rel image, main_part(doc)

          rel_part = reload doc do |saved|
            main_part(saved).get_relationship_by_id rel_id
          end

          rel_part.type.should == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
        end

        it 'adds Relationship tag for image' do
          # here it doesn't really matter what part the image is added to, as long as it shows up in the saved zip/xml
          part = main_part(doc)
          rel_id, image_part = doc.add_image_part_rel image, part

          # now there should be a rel from main_part(doc) to the image part
          rel_part = part.get_relationship_by_id rel_id
          image_rel_path = %<//*:Relationship[@Id = '#{rel_id}'][@Target = '#{rel_part.target_name}']>

          # zip file does not yet contain the new image
          doc.ztree.nxpath(image_rel_path).size.should == 0
          # in-memory structure contains the new image
          doc.xtree.nxpath(image_rel_path).size.should == 1

          target_part = reload doc do |saved|
            # fetch part by relId and verify
            source_part = main_part(saved)
            rel = source_part.get_relationship_by_id rel_id
            rel.target_part
          end

          # outside block so we know the block actually executed
          target_part.name.should == image_part.name
        end

        it 'adds media file for image' do
          # it doesn't really matter what part the image is added to, as long as it shows up in the saved zip/xml
          part = main_part(doc)
          rel_id, image_part = doc.add_image_part_rel image, part

          saved_image_part = reload_document doc do |saved|
            saved.parts[image_part.name]
          end

          # outside block so we know the block actually executed
          saved_image_part.should be_a(Office::ImagePart)
        end

        it 'adds image part' do
          # it doesn't really matter what part the image is added to, as long as it shows up in the saved zip/xml
          part = main_part(doc)
          rel_id, image_part = doc.add_image_part_rel image, part

          image_part = reload_document doc do |saved|
            # fetch part by relId and verify
            source_part = main_part(saved)
            rel = source_part.get_relationship_by_id rel_id
            rel.target_part.name.should == image_part.name
            rel.target_part
          end

          # outside block so we know the block actually executed
          image_part.should be_a(Office::ImagePart)
        end

        describe '#to_data' do
          it 'with value filename' do
            doc.to_data.should be_a(String)
            doc.to_data.should_not be_empty
            doc.to_data.encoding.should == Encoding::ASCII_8BIT
          end

          it 'with nil filename' do
            doc.instance_variable_get(:@filename).should_not be_nil
            doc.remove_instance_variable :@filename

            doc.to_data.should be_a(String)
            doc.to_data.should_not be_empty
            doc.to_data.encoding.should == Encoding::ASCII_8BIT
          end
        end
      end
    end
  end
end
