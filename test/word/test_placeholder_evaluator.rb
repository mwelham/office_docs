#encoding: UTF-8

require 'test/unit'
require 'date'
require 'office/word/placeholder_evaluator'
require 'equivalent-xml'
require 'pry'

class PlaceholderEvaluatorTest < Test::Unit::TestCase
  def test_parse_field_filter
    placeholder = Word::GroupPlaceholder.new("test_field", {:a => :b}, "filter_fields: [a, b, c, d: [e,f: [x,y,z],g], x]")
    filter_field_option = placeholder.field_options.first
    assert filter_field_option.parsed_field_filter == ["a", "b", "c", {"d"=>["e", {"f"=>["x", "y", "z"]}, "g"]}, "x"]


    placeholder = Word::GroupPlaceholder.new("test_field", {:a => :b}, "filter_fields:    [a,b,c,  d:   [e, g,  f:  [x,y,z]],  x]")
    filter_field_option = placeholder.field_options.first
    assert filter_field_option.parsed_field_filter == ["a", "b", "c", {"d"=>["e", "g", {"f"=>["x", "y", "z"]}]}, "x"]
  end

  def test_apply_field_filter_to_group_fields
    field_value = {'a' => '1', 'b' => [{'c' => 3, 'd' => 4, 'x' => 5}, {'c' => 13, 'd' => 14, 'x' => 15}], 'e' => 6, 'y' => 8, 'z' => 9}
    options = "filter_fields: [a, b: [c,d], e: [p], f]"

    placeholder = Word::GroupPlaceholder.new("test_field", field_value, options)

    filter_field_option = placeholder.field_options.first
    filter_field_option.apply_option
    assert placeholder.field_value == [{"a"=>"1", "b"=>[{"c"=>3, "d"=>4}, {"c"=>13, "d"=>14}], "e"=>6}]
  end

  def test_apply_field_filter_to_group_fields_for_blank_filter
    evaluator = Word::PlaceholderEvaluator.new({})
    field_value = {'a' => '1', 'b' => [{'c' => 3, 'd' => 4, 'x' => 5}, {'c' => 13, 'd' => 14, 'x' => 15}], 'e' => 6, 'y' => 8, 'z' => 9}
    options = ["filter_fields:", "filter_fields: []"]

    options.each do |option|
      placeholder = Word::GroupPlaceholder.new("test_field", field_value, option)
      filter_field_option = placeholder.field_options.first
      filter_field_option.apply_option
      assert placeholder.field_value == [{}]
    end
  end

  def test_get_group_list_replacement
    field_value = [{'a' => '1', 'b' => [{'c' => 3, 'd' => 4, 'x' => 5}, {'c' => 13, 'd' => 14, 'x' => 15}], 'e' => 6, 'y' => 8, 'z' => 9}]
    options = "list, filter_fields: [a, b: [c,d], e: [p], f]"

    placeholder = Word::GroupPlaceholder.new("test", field_value, options)

    assert placeholder.replacement == ["test.a\t\t\t1", "test.b", ["-\t\t\ttest.b.c\t\t\t3", "-\t\t\ttest.b.d\t\t\t4", "", "-\t\t\ttest.b.c\t\t\t13", "-\t\t\ttest.b.d\t\t\t14"], "test.e\t\t\t6"]
  end

  def test_bad_options
    field_value = [{'a' => '1', 'b' => [{'c' => 3, 'd' => 4, 'x' => 5}, {'c' => 13, 'd' => 14, 'x' => 15}], 'e' => 6, 'y' => 8, 'z' => 9}]
    options = ["list", "weo: test", "filter_fields: [a, b: [c,d], e: [p], f]", "test", "other_option: heh", "wah: [lol, yo]"]

    begin
    placeholder = Word::GroupPlaceholder.new("test_field", field_value, options.join(','))
    rescue => e
      assert e.message == "Unknown option weo used in the placeholder for test_field."
    end
  end

  def test_split_options_on_commas
    evaluator = Word::Placeholder.new('test',{},'')

    options = "list, weo: test, filter_fields: [a, b: [c,d], e: [p], f], test, other_option: heh, wah: [lol, yo]"
    results = evaluator.send(:split_options_on_commas, (options))

    assert results == ["list", "weo: test", "filter_fields: [a, b: [c,d], e: [p], f]", "test", "other_option: heh", "wah: [lol, yo]"]
  end

  def test_change_multi_select_to_array
    placeholder = Word::Placeholder.new('test',"one,two,three",'each_answer_on_new_line')
    results = placeholder.replacement
    assert results == ["one", "two", "three"]
  end
end
