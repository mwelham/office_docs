require_relative '../lib/office/excel/location'
require_relative '../lib/office/excel/range'

require_relative 'spec_helper'

# copy of the minitest test cases, because specs are easier to zero in on
describe Office::Range do
  let :sampler do (25..512).to_a end

  subject do
    cols = sampler.sample(2).sort
    rows = sampler.sample(2).sort

    tl = Office::Location[ cols.first, rows.first ]
    br = Office::Location[ cols.last, rows.last ]

    described_class.new tl, br
  end

  let :inside do
    col = (subject.top_left.coli..subject.bot_rite.coli).to_a.sample
    row = (subject.top_left.rowi..subject.bot_rite.rowi).to_a.sample
    Office::Location[col,row]
  end

  describe '#cover?' do
    it 'true for locations inside' do
      subject.should cover(inside)
    end

    it 'true for corners' do
      subject.should cover(subject.top_left)
      subject.should cover(subject.bot_rite)
      subject.should cover(subject.bot_left)

      top_rite = Office::Location[subject.bot_rite.coli, subject.top_left.rowi]
      subject.should cover(top_rite)
    end

    it 'true for boundaries' do
      top = Office::Location[(subject.top_left.coli..subject.bot_rite.coli).to_a.sample,subject.top_left.rowi]
      subject.should cover(top)

      bot = Office::Location[(subject.top_left.coli..subject.bot_rite.coli).to_a.sample,subject.bot_rite.rowi]
      subject.should cover(bot)

      lft = Office::Location[subject.top_left.coli,(subject.top_left.rowi..subject.bot_rite.rowi).to_a.sample]
      subject.should cover(lft)

      rgh = Office::Location[subject.bot_rite.coli,(subject.top_left.rowi..subject.bot_rite.rowi).to_a.sample]
      subject.should cover(rgh)
    end

    it 'false for locations outside' do
      subject.should_not cover(subject.bot_rite + [1,0])
      subject.should_not cover(subject.bot_rite + [0,1])
      subject.should_not cover(subject.bot_rite + [1,1])

      subject.should_not cover(subject.top_left - [1,0])
      subject.should_not cover(subject.top_left - [0,1])
      subject.should_not cover(subject.top_left - [1,1])
    end
  end

  describe '.new' do
    it '#width' do
      subject.width.should == subject.bot_rite.coli - subject.top_left.coli + 1
    end

    it '#height' do
      subject.height.should == subject.bot_rite.rowi - subject.top_left.rowi + 1
    end
  end

  describe '#each_by_row' do
    # protection against accidental copypasta in specs
    before :each do
      subject.define_singleton_method(:each_by_col){raise "These are not the methods you are looking for..."}
    end

    it 'not each_by_col' do
      ->{subject.each_by_col}.should raise_error(/not the method/)
    end

    it 'Enumerable without block' do
      subject.each_by_row.should be_a(Enumerable)
    end

    it 'all cells by row' do
      rv = subject.each_by_row.to_a

      # size is correct (maybe unnecessary)
      rv.size.should == subject.width * subject.height

      # corners are correct (maybe unnecessary)
      rv.first.should == subject.top_left
      rv.last.should == subject.bot_rite

      # ok now test that the ordering by row is correct
      first_row = rv.first(subject.width)
      first_row.map(&:rowi).uniq.should == [subject.top_left.rowi]
      first_row.map(&:coli).uniq.should == Array(subject.top_left.coli..subject.bot_rite.coli)
    end

    it 'for single cell' do
      range = Office::Range.new subject.top_left, subject.top_left
      range.each_by_row.to_a.should == [subject.top_left]
    end
  end

  describe '#each_by_col' do
    # protection against accidental copypasta in specs
    before :each do
      subject.define_singleton_method(:each_by_row){raise "These are not the methods you are looking for..."}
    end

    it 'not each_by_row' do
      ->{subject.each_by_row}.should raise_error(/not the method/)
    end

    it 'Enumerable without block' do
      subject.each_by_col.should be_a(Enumerable)
    end

    it 'all cells by col' do
      rv = subject.each_by_col.to_a

      # size is correct (maybe unnecessary)
      rv.size.should == subject.width * subject.height

      # corners are correct (maybe unnecessary)
      rv.first.should == subject.top_left
      rv.last.should == subject.bot_rite

      # ok now test that the ordering by row is correct
      first_col = rv.first(subject.height)
      first_col.map(&:coli).uniq.should == [subject.top_left.coli]
      first_col.map(&:rowi).uniq.should == Array(subject.top_left.rowi..subject.bot_rite.rowi)
    end

    it 'for single cell' do
      range = Office::Range.new subject.top_left, subject.top_left
      range.each_by_col.to_a.should == [subject.top_left]
    end
  end

  describe '#row_of' do
    it 'height is one' do
      row_range = subject.row_of(inside)
      row_range.height.should == 1
    end

    it 'cols are same' do
      row_range = subject.row_of(inside)
      row_range.top_left.coli.should == subject.top_left.coli
      row_range.bot_rite.coli.should == subject.bot_rite.coli
    end

    # a bit redundant, since cols match
    it 'width is same' do
      row_range = subject.row_of(inside)
      row_range.width.should == subject.width
    end

    it 'outside projects location onto range width' do
      outside = subject.bot_rite + [sampler.sample,sampler.sample]
      subject.should_not cover(outside)
      row_range = subject.row_of(outside)

      # row is still outside
      row_range.top_left.rowi.should == row_range.bot_rite.rowi
      subject.should_not cover(row_range.top_left)
      subject.should_not cover(row_range.bot_rite)

      # columns match
      row_range.top_left.coli.should == subject.top_left.coli
      row_range.bot_rite.coli.should == subject.bot_rite.coli
    end
  end
end
