require 'rmagick'
require 'base64'

require_relative 'spec_helper'
require_relative 'modules'
require_relative '../lib/office/nokogiri_extensions'
require_relative '../lib/office/excel/image_drawing'
require_relative '../lib/office/excel/location'

module Office
  describe ImageDrawing do
    MockLoc = Struct.new :coli, :rowi
    MockImage = Struct.new :columns, :rows, :base_filename

    let :loc do MockLoc.new rand(0..257), rand(0..257) end
    let :image do MockImage.new 472, 513, 'the/right/thing.png' end
    let :rel_id do 'sum_dum_rel_id' end
    let :drawing_xdoc do ImageDrawing.build_wsdr.doc end
    let :wsdr do drawing_xdoc.nxpath('*:wsDr').first end

    describe 'with mocks' do
      it 'normal drawing' do
        imdr = ImageDrawing.new img: image, loc: loc, rel_id: rel_id
        imdr.width.should == image.columns * ImageDrawing::PIXELS_TO_EMUS
        imdr.height.should == image.rows * ImageDrawing::PIXELS_TO_EMUS
        imdr.name.should == 'thing.png'

        # verify xml
        imdr.build_anchor wsdr
        wsdr.nxpath('//*:from/*:col').text.should == loc.coli.to_s
        wsdr.nxpath('//*:from/*:row').text.should == loc.rowi.to_s
        wsdr.nxpath('//*:ext/@cx').text.should == (image.columns * ImageDrawing::PIXELS_TO_EMUS).to_s
        wsdr.nxpath('//*:ext/@cy').text.should == (image.rows * ImageDrawing::PIXELS_TO_EMUS).to_s
        wsdr.nxpath('//*:blip/@r:embed').text.should == 'sum_dum_rel_id'
        wsdr.nxpath('//*:nvPicPr/*:cNvPr/@name').text.should == 'thing.png'
      end

      describe 'optional extent' do
        it 'defaults to image size' do
          imdr = ImageDrawing.new img: image, loc: loc, rel_id: rel_id
          imdr.width.should == image.columns * ImageDrawing::PIXELS_TO_EMUS
          imdr.height.should == image.rows * ImageDrawing::PIXELS_TO_EMUS

          # verify xml
          imdr.build_anchor wsdr
          wsdr.nxpath('//*:ext/@cx').text.should == (image.columns * ImageDrawing::PIXELS_TO_EMUS).to_s
          wsdr.nxpath('//*:ext/@cy').text.should == (image.rows * ImageDrawing::PIXELS_TO_EMUS).to_s
        end

        it 'overrides with hash' do
          imdr = ImageDrawing.new img: image, loc: loc, rel_id: 'sum_dum_rel_id', extent: {width: 132, height: 625}
          imdr.width.should == 132 * ImageDrawing::PIXELS_TO_EMUS
          imdr.height.should == 625 * ImageDrawing::PIXELS_TO_EMUS

          # verify xml
          imdr.build_anchor wsdr
          wsdr.nxpath('//*:ext/@cx').text.should == (132 * ImageDrawing::PIXELS_TO_EMUS).to_s
          wsdr.nxpath('//*:ext/@cy').text.should == (625 * ImageDrawing::PIXELS_TO_EMUS).to_s
        end

        it 'overrides with object' do
          imdr = ImageDrawing.new img: image, loc: loc, rel_id: 'sum_dum_rel_id', extent: OpenStruct.new(width: 1066, height: 403)
          imdr.width.should == 1066 * ImageDrawing::PIXELS_TO_EMUS
          imdr.height.should == 403 * ImageDrawing::PIXELS_TO_EMUS

          # verify xml
          imdr.build_anchor(wsdr)
          wsdr.nxpath('//*:ext/@cx').text.should == (1066 * ImageDrawing::PIXELS_TO_EMUS).to_s
          wsdr.nxpath('//*:ext/@cy').text.should == (403 * ImageDrawing::PIXELS_TO_EMUS).to_s
        end
      end

      describe 'drawing.xml' do
        let :loc do MockLoc.new 0, 2 end
        let :image do Magick::ImageList.new FixtureFiles::Image::TEST_IMAGE end
        let :rel_id do 'rId2' end
        let :imdr do ImageDrawing.new img: image, loc: loc, rel_id: rel_id, extent: {width: 28, height: 21} end

        it 'matches the fixture' do
          # because it has MacOS line endings, and then remove the first comment line
          expected_xml = File.read(FixtureFiles::Xml::DRAWING).gsub("\r", "\n").sub(/<!--.*?-->\n*/,'')
          imdr.build_anchor(wsdr).to_xml.should == expected_xml
        end
      end
    end

    describe 'with reals' do
      let :loc do Office::Location.new 'C19' end
      let :image do Magick::ImageList.new FixtureFiles::Image::TEST_IMAGE end
      let :rel_id do Base64.urlsafe_encode64 "#{Time.now}/#{Thread::current.__id__}" end

      it 'normal drawing' do
        imdr = ImageDrawing.new img: image, loc: loc, rel_id: rel_id, extent: {width: 57, height: 91}
        imdr.width.should == 57 * ImageDrawing::PIXELS_TO_EMUS
        imdr.height.should == 91 * ImageDrawing::PIXELS_TO_EMUS
        imdr.name.should == 'test_image.jpg'

        # verify xml
        imdr.build_anchor wsdr
        wsdr.nxpath('//*:from/*:col').text.should == loc.coli.to_s
        wsdr.nxpath('//*:from/*:row').text.should == loc.rowi.to_s
        wsdr.nxpath('//*:ext/@cx').text.should == (57 * ImageDrawing::PIXELS_TO_EMUS).to_s
        wsdr.nxpath('//*:ext/@cy').text.should == (91 * ImageDrawing::PIXELS_TO_EMUS).to_s
        wsdr.nxpath('//*:blip/@r:embed').text.should == rel_id
        wsdr.nxpath('//*:nvPicPr/*:cNvPr/@name').text.should == 'test_image.jpg'
      end
    end

    it '*:cNvPr/@id uniq values' do
      # 'first id is 1'
      imdr = ImageDrawing.new(img: image, loc: loc, rel_id: rel_id, extent: {width: 57, height: 91})
      imdr.build_anchor wsdr
      wsdr.nxpath('//*:cNvPr/@*:id').map(&:text).should == %w[1]

      # 'second id is 2'
      imdr = ImageDrawing.new(img: image, loc: loc, rel_id: rel_id, extent: {width: 22, height: 14})
      imdr.build_anchor wsdr
      wsdr.nxpath('//*:cNvPr/@*:id').map(&:text).should == %w[1 2]

      # 'third id is 3'
      imdr = ImageDrawing.new(img: image, loc: loc, rel_id: rel_id, extent: {width: 57, height: 91})
      imdr.build_anchor wsdr
      wsdr.nxpath('//*:cNvPr/@*:id').map(&:text).should == %w[1 2 3]
    end
  end
end
