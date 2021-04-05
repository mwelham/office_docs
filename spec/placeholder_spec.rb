require_relative 'spec_helper'

require_relative '../lib/office/excel/placeholder_grammar.rb'

module Office
  describe PlaceholderGrammar do
    describe '#read_tokens' do
      def cuddle tokens
        [
          [?{, ?{], [?{, ?{], *tokens, [?}, ?}], [?}, ?}]
        ]
      end
      # tokens are NUMBER IDENTIFIER QUOTE STRING QUOTE BOOLEAN
      it 'single field' do
        tokens = [[:IDENTIFIER, 'some_group']]

        # yyparse is little complicated, but useful. It does not use
        # Racc::Parser#next_token, instead it gets tokens from any iterator.
        # For example, yyparse(obj, :scan) causes calling +obj#scan+, and you
        # can return tokens by yielding them from +obj#scan+.
        #
        # for some reason this fails with the EndOfToken error
        # maybe pass the enum direct with :each ?
        # rv = subject.yyparse enum, :next

        rv = subject.read_tokens cuddle tokens
        subject.field_path.should == %w[some_group]
      end

      it 'field path' do
        tokens = [
          [:IDENTIFIER, 'some_group'],
          [?., ?.],
          [:IDENTIFIER, 'level'],
        ]

        subject.read_tokens cuddle tokens
        subject.field_path.should == %w[some_group level]
      end

      it 'field path with extent' do
        tokens = [
          [:IDENTIFIER, 'some_group'],
          [?., ?.],
          [:IDENTIFIER, 'level'],
          ['|', '|'],
          [:NUMBER, 200],
          ['x', 'x'],
          [:NUMBER, 100],
        ]
        subject.read_tokens cuddle tokens
        subject.field_path.should == %w[some_group level]
        subject.image_extent.should == {width: 200, height: 100}
      end

      it 'field path with bare | is error' do
        tokens = [
          [:IDENTIFIER, 'some_group'],
          [?., ?.],
          [:IDENTIFIER, 'level'],
          ['|', '|'],
        ]

        ->{subject.read_tokens cuddle tokens}.should raise_error(Racc::ParseError)
      end

      it 'field path with keywords' do
        tokens = [
          [:IDENTIFIER, 'some_group'],
          [?., ?.],
          [:IDENTIFIER, 'level'],
          [?|, ?|],
          [:IDENTIFIER, 'show_coordinate_info'],
          [?:, ?:],
          [:false, :false],
        ]
        rv = subject.read_tokens cuddle tokens
        subject.field_path.should == %w[some_group level]
        subject.keywords.should == {show_coordinate_info: false}
      end

      it 'field path with functor' do
        tokens = [
          [:IDENTIFIER, 'some_group'],
          [?., ?.],
          [:IDENTIFIER, 'level'],
          [?|, ?|],
          [:IDENTIFIER, 'layout'],
          [?(, ?(],
          [:RANGE, 'A1:G7'],
          [?), ?)],
        ]
        rv = subject.read_tokens cuddle tokens
        subject.field_path.should == %w[some_group level]
        subject.functors.should == {layout: 'A1:G7'}
      end
    end

    describe '#parse' do
      (Pathname(__dir__) +'fixtures/all-placeholders.txt').each_line do |line|
        it "tokenizes #{line}" do
          tokens = Array(PlaceholderGrammar.tokenize line)
          subject.read_tokens tokens.dup
          ->{subject.read_tokens tokens}.should_not raise_error
        rescue
          binding.pry
        end
      end

      it "tokenizes unquoted keyword values" do
        line = %<{{ submitted_at | date_time_format: %d &m %y, capitalize, separator: ;, justify  }}>
        tokens = Array(PlaceholderGrammar.tokenize line)
        subject.read_tokens tokens
        subject.to_h.should == {
          :field_path=>['submitted_at'],
          :image_extent=>nil,
          :keywords=>{:date_time_format=>"%d &m %y", :capitalize=>true, :separator=>";", :justify=>true},
          :functors=>{}
        }
      end

      it "tokenizes bracketed extent" do
        line = %<{{ fields.Group | image_size: [100x200]}}>
        tokens = Array(PlaceholderGrammar.tokenize line)
        subject.read_tokens tokens
        subject.to_h.should == {:field_path=>["fields", "Group"], :image_extent=>{}, :keywords=>{image_size: {:width=>"100", :height=>"200"}}, :functors=>{}}
      end
    end

    describe '#parse' do
      xit 'image extent' do
        ph = Placeholder.parse 'entries.your_picture|100x200'
        ph.name.should == 'entries.your_picture'
        ph.options.should == {extent: [100,200]}
      end

      xit 'image extent' do
        ph = Placeholder.parse 'entries.your_picture|size(100,200)'
        ph.name.should == 'entries.your_picture'
        ph.options.should == {size: [100,200]}
      end

      xit 'image extent' do
        ph = Placeholder.parse 'entries.entries.group|layout: "d3:g4"'
        ph.name.should == 'entries.entries.group'
        ph.options.should == {layout: %w[d3 g4]}
      end
    end
  end

end
