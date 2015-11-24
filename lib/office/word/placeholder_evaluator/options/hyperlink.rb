module Word
  module Options
    class Hyperlink < Word::Option
      def importance
        20
      end

      def apply_option
        if params.downcase == 'true'
          coord_info = placeholder.field_value[1]

          if !coord_info.nil?
            lat = coord_info.match(/lat=((\d+|-\d)+\.\d+)/)
            long = coord_info.match(/long=((\d+|-\d)+\.\d+)/)
            if !lat.nil? && !long.nil?
              placeholder.render_options[:hyperlink] = "http://maps.google.com/?q=#{lat[1]},#{long[1]}"
            end
          end
        end
      end
    end
  end
end
