require_relative 'spec_helper'

require_relative '../lib/office/excel/placeholder_grammar.rb'
require_relative '../lib/office/excel/placeholder_lexer.rb'

module Office
  describe PlaceholderGrammar do
    describe '#read_tokens' do
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

        rv = subject.read_tokens tokens
        subject.field_path.should == %w[some_group]
      end

      it 'field path' do
        tokens = [
          [:IDENTIFIER, 'some_group'],
          [?., ?.],
          [:IDENTIFIER, 'level'],
        ]

        subject.read_tokens tokens
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
        subject.read_tokens tokens
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

        ->{subject.read_tokens tokens}.should raise_error(Racc::ParseError)
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
        rv = subject.read_tokens tokens
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
        rv = subject.read_tokens tokens
        subject.field_path.should == %w[some_group level]
        subject.functors.should == {layout: 'A1:G7'}
      end
    end

    # NOTE generated from
    #   Word::PlaceholderEvaluator#initialize
    # using
    #   File.open('/tmp/all.txt','a'){|io| io.puts placeholder[:placeholder_text]}
    # then
    #   rake test:all
    describe '#parse' do
      describe 'extracted' do
        (Pathname(__dir__) +'fixtures/all-placeholders.txt').each_line do |line|
          it "tokenizes #{line}" do
            tokens = Array(PlaceholderLexer.tokenize line)
            ->{subject.read_tokens tokens}.should_not raise_error
            # binding.pry if line =~ /show_coordinate_info:/
          end
        end
      end

      describe 'fixtures' do
        it "empty placeholder" do
          tokens = Array(PlaceholderLexer.tokenize "{{}}")
          subject.read_tokens(tokens).should == {
            :field_path=>[],
            :image_extent=>nil,
            :keywords=>{},
            :functors=>{}
          }
        end

        it "tokenizes unquoted keyword values" do
          line = %<{{ submitted_at | date_time_format: %d &m %y, capitalize, separator: ;, justify  }}>
          tokens = Array(PlaceholderLexer.tokenize line)
          subject.read_tokens tokens
          subject.to_h.should == {
            :field_path=>['submitted_at'],
            :image_extent=>nil,
            :keywords=>{:date_time_format=>"%d &m %y", :capitalize=>true, :separator=>";", :justify=>true},
            :functors=>{}
          }
        end

        it 'boolean keyword' do
          line = '{{doh|hedge: true}}'
          tokens = Array(PlaceholderLexer.tokenize line)
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%w[doh], :image_extent=>nil, :keywords=>{hedge: true}, :functors=>{}}
        end

        it 'numeric keyword' do
          line = '{{doh|nuts: 17}}'
          tokens = Array(PlaceholderLexer.tokenize line)
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%w[doh], :image_extent=>nil, :keywords=>{nuts: 17}, :functors=>{}}
        end

        it "bracketed image_size" do
          line = %<{{ fields.Group | image_size: [100x200]}}>
          tokens = Array(PlaceholderLexer.tokenize line)
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%w[fields Group], :image_extent=>nil, :keywords=>{image_size: {:width=>"100", :height=>"200"}}, :functors=>{}}
        end

        it 'image extent' do
          line = '{{entries.your_picture|100x200}}'
          tokens = Array(PlaceholderLexer.tokenize line)
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%w[entries your_picture], :image_extent=>{:width=>"100", :height=>"200"}, :keywords=>{}, :functors=>{}}
        end

        it 'image functor size' do
          line = '{{entries.your_picture|size(100,200)}}'
          tokens = Array(PlaceholderLexer.tokenize line)
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%w[entries your_picture], :image_extent=>nil, :functors=>{size: [100, 200]}, :keywords=>{}}
        end

        it 'multi-value functor' do
          line = '{{doh|stuff(1,2,3,4,5)}}'
          tokens = Array(PlaceholderLexer.tokenize line)
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%w[doh], :image_extent=>nil, :functors=>{stuff: [1,2,3,4,5]}, :keywords=>{}}
        end

        describe "layout" do
          it 'quoted layout keyword' do
            line = '{{entries.entries.group|layout: "d3:g4"}}'
            tokens = Array(PlaceholderLexer.tokenize line)
            subject.read_tokens tokens
            subject.to_h.should == {:field_path=>%w[entries entries group], :image_extent=>nil, :keywords=>{layout: 'd3:g4'}, :functors=>{}}
          end

          it 'bare layout keyword' do
            line = '{{entries.entries.group|layout: d3:g4}}'
            tokens = Array(PlaceholderLexer.tokenize line)
            subject.read_tokens tokens
            subject.to_h.should == {:field_path=>%w[entries entries group], :image_extent=>nil, :keywords=>{layout: 'd3:g4'}, :functors=>{}}
          end

          # breaks because of magic quoting
          it 'layout functor' do
            line = '{{entries.entries.group|layout(d3:g4)}}'
            subject.read_tokens PlaceholderLexer.tokenize line
            subject.to_h.should == {:field_path=>%w[entries entries group], :image_extent=>nil, :functors=>{layout: 'd3:g4'}, :keywords=>{}}
          end
        end
      end
    end

    describe '#parse' do
      it 'implement lex then parse'
    end
  end
end
