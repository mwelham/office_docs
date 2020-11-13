require_relative '../lib/office/excel/location'
require_relative '../lib/office/excel/range'

require_relative 'spec_helper'

# copy of the minitest test cases, because specs are easier to zero in on
describe Office::Location do
  describe '*' do
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
