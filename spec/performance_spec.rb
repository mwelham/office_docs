require_relative 'spec_helper'

require 'date'
require 'benchmark'

# This is a place to hang performance comparisons that don't naturally belong
# somewhere else. See spec_helper.rb config for where performance tests are
# excluded by default.

module AsIs
  # should probably be lambdas in a lookup, or clauses in the case
  def as_decimal(value)
    value.to_f
  end

  def as_integer(value)
    value.to_i
  end

  DATE_TIME_EPOCH = DateTime.new(1900, 1, 1, 0, 0, 0) - 2
  DATE_EPOCH = Date.new(1900, 1, 1) - 2

  def as_datetime(value)
    DATE_TIME_EPOCH + value.to_f
  end

  def as_date(value)
    # This was originally DateTime, and I don't know why it wasn't just Date
    # Date.new(1900, 1, 1, 0, 0, 0) + value.to_i - 2
    DATE_EPOCH + value.to_i
  end

  def as_time(value)
    as_datetime(value).to_time
  end
end

class Splat
  include AsIs
  extend AsIs

  LAMD_HASH = {
     1 => -> unformatted_value {as_integer(unformatted_value)},
     3 => -> unformatted_value {as_integer(unformatted_value)},
     0 => -> unformatted_value {as_decimal(unformatted_value)},
     2 => -> unformatted_value {as_decimal(unformatted_value)},
     4 => -> unformatted_value {as_decimal(unformatted_value)},
     9 => -> unformatted_value {as_decimal(unformatted_value)},
    10 => -> unformatted_value {as_decimal(unformatted_value)},
    11 => -> unformatted_value {as_decimal(unformatted_value)},
    37 => -> unformatted_value {as_decimal(unformatted_value)},
    38 => -> unformatted_value {as_decimal(unformatted_value)},
    39 => -> unformatted_value {as_decimal(unformatted_value)},
    40 => -> unformatted_value {as_decimal(unformatted_value)},
    48 => -> unformatted_value {as_decimal(unformatted_value)},
    15 => -> unformatted_value {as_date(unformatted_value)},
    16 => -> unformatted_value {as_date(unformatted_value)},
    17 => -> unformatted_value {as_date(unformatted_value)},
    22 => -> unformatted_value {as_datetime(unformatted_value)},
    18 => -> unformatted_value {as_time(unformatted_value)},
    19 => -> unformatted_value {as_time(unformatted_value)},
    20 => -> unformatted_value {as_time(unformatted_value)},
    21 => -> unformatted_value {as_time(unformatted_value)},
    45 => -> unformatted_value {as_time(unformatted_value)},
    46 => -> unformatted_value {as_time(unformatted_value)},
    47 => -> unformatted_value {as_time(unformatted_value)},
  }

  # of course this assumes that the method definitions don't change in the meantime.
  METH_HASH = {
     1 => method(:as_integer),
     3 => method(:as_integer),
     0 => method(:as_decimal),
     2 => method(:as_decimal),
     4 => method(:as_decimal),
     9 => method(:as_decimal),
    10 => method(:as_decimal),
    11 => method(:as_decimal),
    37 => method(:as_decimal),
    38 => method(:as_decimal),
    39 => method(:as_decimal),
    40 => method(:as_decimal),
    48 => method(:as_decimal),
    15 => method(:as_date),
    16 => method(:as_date),
    17 => method(:as_date),
    22 => method(:as_datetime),
    18 => method(:as_time),
    19 => method(:as_time),
    20 => method(:as_time),
    21 => method(:as_time),
    45 => method(:as_time),
    46 => method(:as_time),
    47 => method(:as_time),
  }

  SPLAT_HASH = METH_HASH

  # Just a working thing to construct splatcase
  SPLATCASE_ARY = SPLAT_HASH.group_by{|_,v| v.name}.map{|name,ary| [name,ary.map{|(index,_methods)| index}] }.to_h

  def splatcase index, unformatted_value
    case index
    when 1, 3
      # as_integer unformatted_value
      unformatted_value.to_i
    when 0, 2, 4, 9, 10, 11, 37, 38, 39, 40, 48
      # as_decimal unformatted_value
      unformatted_value.to_f
    when 15, 16, 17
      # as_date unformatted_value
      DATE_EPOCH + unformatted_value.to_i
    when 22
      # as_datetime unformatted_value
      DATE_TIME_EPOCH + unformatted_value.to_f
    when 18, 19, 20, 21, 45, 46, 4
      as_time unformatted_value
    end
  end

  def hash index, unformatted_value
    if λ = SPLAT_HASH[index]
      λ[unformatted_value]
    else
      unformatted_value
    end
  end

  WHEN_ARY = SPLAT_HASH.each_with_object [] do |(ix,λ),ary| ary[ix] = λ end

  def array index, unformatted_value
    if λ = WHEN_ARY[index]
      λ[unformatted_value]
    else
      unformatted_value
    end
  end

  def whens index, unformatted_value
    case index
    when 0  #    General
      as_decimal(unformatted_value)
    when 1  #    0
      as_integer(unformatted_value)
    when 2  #    0.00
      as_decimal(unformatted_value)
    when 3  #    #,##0
      as_integer(unformatted_value)
    when 4  #    #,##0.00
      as_decimal(unformatted_value)
    when 9  #    0%
      as_decimal(unformatted_value)
    when 10 #    0.00%
      as_decimal(unformatted_value)
    when 11 #    0.00E+00
      as_decimal(unformatted_value)
    #when 12 #    # ?/?
    #when 13 #    # ??/??
    when 14 #    mm-dd-yy
      as_date(unformatted_value)
    when 15 #    d-mmm-yy
      as_date(unformatted_value)
    when 16 #    d-mmm
      as_date(unformatted_value)
    when 17 #    mmm-yy
      as_date(unformatted_value)
    when 18 #    h:mm AM/PM
      as_time(unformatted_value)
    when 19 #    h:mm:ss AM/PM
      as_time(unformatted_value)
    when 20 #    h:mm
      as_time(unformatted_value)
    when 21 #    h:mm:ss
      as_time(unformatted_value)
    when 22 #    m/d/yy h:mm
      as_datetime(unformatted_value)
    when 37 #    #,##0 ;(#,##0)
      as_decimal(unformatted_value)
    when 38 #    #,##0 ;[Red](#,##0)
      as_decimal(unformatted_value)
    when 39 #    #,##0.00;(#,##0.00)
      as_decimal(unformatted_value)
    when 40 #    #,##0.00;[Red](#,##0.00)
      as_decimal(unformatted_value)
    when 45 #    mm:ss
      as_time(unformatted_value)
    when 46 #    [h]:mm:ss
      as_time(unformatted_value)
    when 47 #    mmss.0
      as_time(unformatted_value)
    when 48 #    ##0.0E+0
      as_decimal(unformatted_value)
    #when 49 #    @
    else
      unformatted_value
    end
  end
end

describe Splat do
  it 'performance', performance: true do
    domain = (0..49).to_a
    repeats = 100_000
    # comparable, although splatcase is steadily the fastest
    #
    #                 user     system      total        real
    # whens       0.067839   0.000191   0.068030 (  0.068027)
    # splatcase   0.063505   0.000193   0.063698 (  0.063697)
    # ary         0.069882   0.000195   0.070077 (  0.070075)
    # hash        0.071936   0.000213   0.072149 (  0.072159)
    Benchmark.bmbm do |results|
      results.report 'whens'     do repeats.times{subject.whens     domain.sample, 31415.27} end
      results.report 'splatcase' do repeats.times{subject.splatcase domain.sample, 31415.27} end
      results.report 'ary'       do repeats.times{subject.array     domain.sample, 31415.27} end
      results.report 'hash'      do repeats.times{subject.hash      domain.sample, 31415.27} end
    end
  end
end
