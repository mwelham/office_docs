require_relative 'spec_helper'

require_relative '../lib/office/excel/placeholder.rb'

module Office
  describe PlaceholderGrammar do
    describe '#read_tokens' do
      # tokens are NUMBER IDENTIFIER QUOTE STRING QUOTE BOOLEAN
      it 'single field' do
        subject.read_tokens [[:IDENTIFIER, 'some_group']]
        subject.field_path.should == %i[some_group]
      end

      it 'field path' do
        tokens = [
          [:IDENTIFIER, 'some_group'],
          [?., ?.],
          [:IDENTIFIER, 'level'],
        ]

        subject.read_tokens tokens
        subject.field_path.should == %i[some_group level]
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
        subject.field_path.should == %i[some_group level]
        subject.image_extent.should == {width: 200, height: 100}
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
        subject.field_path.should == %i[some_group level]
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
        subject.field_path.should == %i[some_group level]
        subject.functors.should == {layout: 'A1:G7'}
      end
    end

    describe 'parse' do
      describe 'extracted', extracted: true do
        # NOTE generated from
        #   Word::PlaceholderEvaluator#initialize
        # using
        #   File.open('<...>/fixtures/all-placeholders.txt','a'){|io| io.puts placeholder[:placeholder_text]}
        # then
        #   rake test:all
        (Pathname(__dir__) +'fixtures/all-placeholders.txt').each_line do |line|
          it "tokenizes #{line}" do
            tokens = PlaceholderLexer.tokenize line
            ->{subject.read_tokens tokens}.should_not raise_error
          end
        end
      end

      describe 'fixtures' do
        it "empty placeholder" do
          tokens = PlaceholderLexer.tokenize "{{}}"
          subject.read_tokens(tokens).should == {
            :field_path=>[],
            :image_extent=>nil,
            :keywords=>{},
            :functors=>{}
          }
        end

        it "really empty placeholder" do
          tokens = PlaceholderLexer.tokenize ""
          subject.read_tokens(tokens).should == {
            :field_path=>[],
            :image_extent=>nil,
            :keywords=>{},
            :functors=>{}
          }
        end

        it "tokenizes unquoted keyword values" do
          line = %<{{ submitted_at | date_time_format: %d &m %y, capitalize, separator: ;, justify  }}>
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {
            :field_path=>[:submitted_at],
            :image_extent=>nil,
            :keywords=>{:date_time_format=>"%d &m %y", :capitalize=>true, :separator=>";", :justify=>true},
            :functors=>{}
          }
        end

        it "fails on field whitespaces" do
          line = %<{{controller.streams[0].the word}>
          tokens = PlaceholderLexer.tokenize line
          ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, /Unexpected word at/)
        end

        it "weird whitespaces" do
          line = %<{{   submitted_at      |  date_time_format: %d   &m  %y,capitalize,separator:;,justify}     }>
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {
            :field_path=>[:submitted_at],
            :image_extent=>nil,
            :keywords=>{:date_time_format=>"%d   &m  %y", :capitalize=>true, :separator=>";", :justify=>true},
            :functors=>{}
          }
        end

        it 'boolean keyword' do
          line = '{{doh|hedge: true}}'
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[doh], :image_extent=>nil, :keywords=>{hedge: true}, :functors=>{}}
        end

        it 'numeric keyword' do
          line = '{{doh|nuts: 17}}'
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[doh], :image_extent=>nil, :keywords=>{nuts: 17}, :functors=>{}}
        end

        it "bracketed image_size" do
          line = %<{{ fields.Group | image_size: [100x200]}}>
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[fields Group], :image_extent=>nil, :keywords=>{image_size: {:width=>100, :height=>200}}, :functors=>{}}
        end

        it 'image extent' do
          line = '{{entries.your_picture|100x200}}'
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[entries your_picture], :image_extent=>{:width=>100, :height=>200}, :keywords=>{}, :functors=>{}}
        end

        it 'spaced extent' do
          line = '{{entries.your_picture| 100  X 200    }}'
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[entries your_picture], :image_extent=>{:width=>100, :height=>200}, :keywords=>{}, :functors=>{}}
        end

        it 'size functor arguments' do
          line = '{{entries.your_picture|size(100,200)}}'
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[entries your_picture], :image_extent=>nil, :functors=>{size: [100, 200]}, :keywords=>{}}
        end

        it 'size functor extent' do
          line = '{{entries.your_picture|extent_size(100x200)}}'
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[entries your_picture], :image_extent=>nil, :functors=>{extent_size: {height: 200, width: 100}}, :keywords=>{}}
        end

        it 'one space functor extent' do
          line = '{{entries.your_picture|extent_size( 100 x 200 )}}'
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[entries your_picture], :image_extent=>nil, :functors=>{extent_size: {height: 200, width: 100}}, :keywords=>{}}
        end

        it 'spaced functor extent' do
          line = "{{entries.your_picture|extent_size(  100  x  \n200  )}}"
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[entries your_picture], :image_extent=>nil, :functors=>{extent_size: {height: 200, width: 100}}, :keywords=>{}}
        end

        it 'multi-value functor' do
          line = '{{doh|stuff(1,2,3,4,5)}}'
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {:field_path=>%i[doh], :image_extent=>nil, :functors=>{stuff: [1,2,3,4,5]}, :keywords=>{}}
        end

        it 'everything' do
          line = %<                  group.where._whatever[0].classes[3].full_name |justify,transition: ;,date_format("%d-%b-%y"),layout(aaf4:aag7),neutralise: 'ph', ph: 10,froomative(true),negatory(15)>
          tokens = PlaceholderLexer.tokenize line
          subject.read_tokens tokens
          subject.to_h.should == {
            :field_path=>[:group, :where, :_whatever, 0, :classes, 3, :full_name],
            :image_extent=>nil,
            :keywords=>{:justify=>true, :transition=>";", :neutralise=>"ph", :ph=>10},
            :functors=>{:date_format=>"%d-%b-%y", :layout=>"aaf4:aag7", froomative: true, negatory: 15}
          }
        end

        describe "layout" do
          it 'quoted layout keyword' do
            line = '{{entries.entries.group|layout: "d3:g4"}}'
            tokens = PlaceholderLexer.tokenize line
            subject.read_tokens tokens
            subject.to_h.should == {:field_path=>%i[entries entries group], :image_extent=>nil, :keywords=>{layout: 'd3:g4'}, :functors=>{}}
          end

          it 'bare layout keyword' do
            line = '{{entries.entries.group|layout: d3:g4}}'
            tokens = PlaceholderLexer.tokenize line
            subject.read_tokens tokens
            subject.to_h.should == {:field_path=>%i[entries entries group], :image_extent=>nil, :keywords=>{layout: 'd3:g4'}, :functors=>{}}
          end

          # breaks because of magic quoting
          it 'layout functor' do
            line = '{{entries.entries.group|layout(d3:g4)}}'
            subject.read_tokens PlaceholderLexer.tokenize(line)
            subject.to_h.should == {:field_path=>%i[entries entries group], :image_extent=>nil, :functors=>{layout: 'd3:g4'}, :keywords=>{}}
          end
        end

        describe 'failures' do
          it 'first half placeholder' do
            tokens = PlaceholderLexer.tokenize "{{"
            ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, 'Unexpected end after {{')
          end

          it 'second half placeholder' do
            tokens = PlaceholderLexer.tokenize "}}"
            ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, "Unexpected } at 0:0 in '}'")
          end

          it 'field path trailing .' do
            tokens = PlaceholderLexer.tokenize "some_group.level."
            ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, 'Unexpected end after some_group.level.')
          end

          it 'field path with bare |' do
            tokens = PlaceholderLexer.tokenize "some_group.level|"
            ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, 'Unexpected end after some_group.level|')
          end

          it 'keyword:' do
            tokens = PlaceholderLexer.tokenize "some_group.level|translate:"
            ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, 'Unexpected end after some_group.level|translate:')
          end

          it 'unmatched quote' do
            tokens = PlaceholderLexer.tokenize "some_group.level|translate: 'this thing"
            ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, "Unexpected this at 0:29 in 'some_group.level|translate: 'this'")
          end

          it 'bad range in layout' do
            tokens = PlaceholderLexer.tokenize "some_group.level|translate: 12,layout(a15)"
            ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, "Unexpected ) at 0:41 in 'some_group.level|translate: 12,layout(a15)'")
          end

          it 'numerics with trailing puntuation' do
            tokens = PlaceholderLexer.tokenize "some_group.level|translate: 12,style(1):"
            ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, "Unexpected : at 0:39 in 'some_group.level|translate: 12,style(1):'")
          end

          it 'naked value functor' do
            line = '{{doh|date_format(%d-%b-%y)}}'
            tokens = PlaceholderLexer.tokenize line
            ->{subject.read_tokens tokens}.should raise_error(Office::PlaceholderGrammar::ParseError, "Unexpected % at 0:18 in '{{doh|date_format(%'")
          end
        end
      end
    end
  end

  describe Placeholder do
    describe '.parse' do
      it 'fairly complex' do
        line = %<{{ group.where._whatever[0].classes[3].full_name |justify,transition: ;,  320X200   ,date_format("%d-%b-%y"),layout(aaf4:aag7),neutralise: 'ph', ph: 10,froomative(true),negatory(15)}}>
        ph = Placeholder.parse line

        ph.field_path.should == [:group, :where, :_whatever, 0, :classes, 3, :full_name]
        ph.options.should == {
          :justify=>true,
          :transition=>";",
          :neutralise=>"ph",
          :ph=>10,
          :date_format=>"%d-%b-%y",
          :layout=>"aaf4:aag7",
          :froomative=>true,
          :negatory=>15}
      end

      it 'fails and raises' do
        line = %<{{ group.where._whatever[0].classes[3].full name |justify,transition: ;,  320X200   ,date_format("%d-%b-%y"),layout(aaf4:aag7),neutralise: 'ph', ph: 10,froomative(true),negatory(15)}}>
        ->{Placeholder.parse line}.should raise_error(Office::PlaceholderGrammar::ParseError, /Unexpected name/)
      end
    end
  end
end

