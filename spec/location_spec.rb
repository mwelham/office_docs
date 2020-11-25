require_relative '../lib/office/excel/location'
require_relative '../lib/office/excel/range'

require_relative 'spec_helper'

# copy of the minitest test cases, because specs are easier to zero in on
describe Office::Location do
  describe 'infinite' do
    it 'largest' do
      Office::Location.largest.to_a.should == [Float::INFINITY, Float::INFINITY]
      Office::Location.largest.to_s.should == '+∞'
      Office::Location.largest.clone.to_s.should == '+∞'

      ->{Office::Location.largest.dup.to_s}.should raise_error(/can't dup.*largest/)
    end

    it 'smallest' do
      Office::Location.smallest.to_a.should == [-Float::INFINITY, -Float::INFINITY]
      Office::Location.smallest.to_s.should == '-∞'
      Office::Location.smallest.clone.to_s.should == '-∞'

      # this isn't really proper, but oh well
      ->{Office::Location.smallest.dup}.should raise_error(/can't dup.*smallest/)
    end
  end

  describe 'operator *' do
    it 'extends by some' do
      loc = Office::Location.new 'L2'
      range = loc * [4,3]

      range.width.should == 4
      range.height.should == 3

      range.bot_rite.should == 'O4'
    end

    it 'extends by 1' do
      loc = Office::Location.new 'L2'
      range = loc * [1,1]

      range.width.should == 1
      range.height.should == 1

      range.bot_rite.should == 'L2'
    end
  end
end
