module VersionCompatibility
  if RUBY_VERSION < '2.7.0'
    refine Time do
      def floor
        Time.new(year, month, day, hour, min, sec.floor, utc_offset)
      end
    end
  end

  if RUBY_VERSION < '2.5.0'
    refine Hash do
      def transform_keys &blk
        return enum_for __method__ unless block_given?
        each_with_object Hash.new do |(key,value),ha|
          ha[blk[key]] = value
        end
      end
    end
  end
end

