require 'ostruct'
require 'nokogiri'

module Office
  class ImageDrawing
    # DrawingML use EMUs which are English Metric Units.
    PIXELS_TO_EMUS = 9525

    NAMESPACE_DECLS = {
      :'xmlns:xdr'   => 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
      :'xmlns:a'     => 'http://schemas.openxmlformats.org/drawingml/2006/main',
      :'xmlns:r'     => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }

    # img is something that understands (or really columns:Pixel, rows:Pixel and base_filename:String), eg ImageMagick
    # loc is something that understands rowi and coli as 0-based indices, eg Office::Location
    # rel_id is a string containing the relationship id to the image resource displayed by this drawing.
    # extent is optional {width: 57, height: 91} in pixels
    def initialize img:, loc:, rel_id:, extent: nil
      @img, @loc, @rel_id  = img, loc, rel_id
      @extent = self.class.dotify extent
    end

    # make extent hashes dottable
    def self.dotify extent
      case extent
      when Hash
        OpenStruct.new(extent).freeze
      else
        extent
      end
    end

    def width; (@extent&.width || @img.columns) * PIXELS_TO_EMUS end
    def height; (@extent&.height || @img.rows) * PIXELS_TO_EMUS end
    def col; @loc.coli end
    def row; @loc.rowi end
    def name; File.basename @img.base_filename end
    def rel_id; @rel_id end

    def largest_cnvpr_id wsdr_node
      wsdr_node.nxpath('//*:cNvPr/@*:id').map(&:text).max&.to_i || 0
    end

    # Append a oneCellAnchor tag to the given wsDr tag (which should be in a
    # drawing.xml properly reference from the sheet).
    #
    # Returns the updated wsdr_node.
    def build_anchor wsdr_node
      r_ns = NAMESPACE_DECLS.slice :'xmlns:r'
      Nokogiri::XML::Builder.with wsdr_node do |bld|
        anchor_builder = bld.oneCellAnchor do
          bld.from do
            bld.col col
            bld.colOff 0
            bld.row row
            bld.rowOff 0
          end
          bld.ext cx: width, cy: height
          bld.pic do
            bld.nvPicPr do
              # this can also have title: for a caption. But it doesn't seem to show up for xlsx files.
              # id: must be unique in this document. There's only one image for this drawing, so this complies.
              bld.cNvPr id: (largest_cnvpr_id(wsdr_node)+1), name: name
              bld.cNvPicPr preferRelativeResize: 0
            end
            bld.blipFill do
              # NOTE blip must have the namespace declaration. Excel over-optimises the wsDr
              # tag which quite often doesn't contain the r namespace.
              bld[:a].blip **r_ns, cstate: 'print', 'r:embed': rel_id
              bld[:a].stretch do
                bld[:a].fillRect
              end
            end
            bld.spPr do
              bld[:a].prstGeom prst: 'rect' do
                bld[:a].avLst
              end
              bld[:a].noFill
            end
          end
          bld.clientData fLocksWithSheet: 0
        end
      end.doc
    end

    def self.build_wsdr
      # apparently this is the only way to get standalone
      # because to_xml(encoding: "UTF-8" standalone: "yes") leaves out the standalone
      # TODO verify that - maybe one of the write_to or to_xml options works https://github.com/sparklemotion/nokogiri/wiki/Cheat-sheet#working-with-a-nokogirixmlnode
      xml_decl = Nokogiri::XML %|<?xml version="1.0" encoding="UTF-8" standalone="yes"?>|

      # TODO this generates LF only - will that work ok with MacOS and Windows?
      # builder code was fine in the sheet and cell code so probably will be fine.
      # we only need the namespaces we refer to
      Nokogiri::XML::Builder.with xml_decl do |bld|
        bld[:xdr].wsDr NAMESPACE_DECLS
      end # builder
    end
  end
end
