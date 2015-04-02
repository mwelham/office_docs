#encoding: UTF-8

require 'test/unit'
require 'date'
require 'office/word/placeholder_evaluator'
require 'equivalent-xml'
require 'pry'

class PlaceholderEvaluatorTest < Test::Unit::TestCase
  def test_parse_field_filter
    evaluator = Word::PlaceholderEvaluator.new({})
    result = evaluator.parse_field_filter("fields: [a, b, c, d: [e,f: [x,y,z],g], x]")
    assert result == ["a", "b", "c", {"d"=>["e", {"f"=>["x", "y", "z"]}, "g"]}, "x"]

    result = evaluator.parse_field_filter("fields:    [a,b,c,  d:   [e, g,  f:  [x,y,z]],  x]")
    assert result == ["a", "b", "c", {"d"=>["e", "g", {"f"=>["x", "y", "z"]}]}, "x"]
  end

  def test_apply_field_filter_to_group_fields
    evaluator = Word::PlaceholderEvaluator.new({})
    field_value = {'a' => '1', 'b' => [{'c' => 3, 'd' => 4, 'x' => 5}, {'c' => 13, 'd' => 14, 'x' => 15}], 'e' => 6, 'y' => 8, 'z' => 9}
    result = evaluator.apply_field_filter_to_group_fields("fields: [a, b: [c,d], e: [p], f", field_value)
    assert result == {"a"=>"1", "b"=>[{"c"=>3, "d"=>4}, {"c"=>13, "d"=>14}], "e"=>6}
  end

  def test_apply_field_filter_to_group_fields_for_blank_filter
    evaluator = Word::PlaceholderEvaluator.new({})
    field_value = {'a' => '1', 'b' => [{'c' => 3, 'd' => 4, 'x' => 5}, {'c' => 13, 'd' => 14, 'x' => 15}], 'e' => 6, 'y' => 8, 'z' => 9}
    result = evaluator.apply_field_filter_to_group_fields("fields:", field_value)
    assert result == {}

    result = evaluator.apply_field_filter_to_group_fields("fields: []", field_value)
    assert result == {}
  end

  def test_apply_options_to_other_value
    field_value = [{'a' => '1', 'b' => [{'c' => 3, 'd' => 4, 'x' => 5}, {'c' => 13, 'd' => 14, 'x' => 15}], 'e' => 6, 'y' => 8, 'z' => 9}]
    options = "list, weo: test, fields: [a, b: [c,d], e: [p], f], test, other_option: heh, wah: [lol, yo]"
    evaluator = Word::PlaceholderEvaluator.new({})
    result = evaluator.apply_options_to_group_value('test', field_value, options)
    assert result == ["test.a\t\t\t1", "test.b", ["-\t\t\ttest.b.c\t\t\t3", "-\t\t\ttest.b.d\t\t\t4", "", "-\t\t\ttest.b.c\t\t\t13", "-\t\t\ttest.b.d\t\t\t14"], "test.e\t\t\t6"]
  end
end
