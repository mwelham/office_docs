require_relative '../lib/office/excel/location'
require_relative '../lib/office/excel/range'

require_relative 'spec_helper'

# copy of the minitest test cases, because specs are easier to zero in on
describe Office::Range do
  # TODO this is a bit dim :-|
  def sample sampler_fn, conditional_fn = ->_{true}
    loop do
      val = sampler_fn[]
      break val if conditional_fn[val]
    end
  end

  subject do
    sampler_fn = ->{rand 256..512}

    tl = Office::Location[ (fst_coli = sample(sampler_fn)), (fst_rowi = sample(sampler_fn)) ]

    # NOTE this would cause an infinite loop if _ fst_xxx value happens to be the last in the range.
    br = Office::Location[ sample(sampler_fn, ->v{v >= fst_coli}), sample(sampler_fn, ->v{v >= fst_rowi}) ]
    described_class.new tl, br
  end

  describe '#cover?' do
    it 'true for locations inside'
    it 'true for corners'
    it 'true for boundaries'
    it 'false for locations outside'
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
end
