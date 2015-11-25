module Word
  module Options
    class ShowCoordinateInfo < Word::Option
      def importance
        5
      end

      def apply_option

        if params.downcase == 'false' && placeholder.field_value.length > 1
          placeholder.field_value = [placeholder.field_value[0]]
        elsif params != 'true'
          # e.g. %lat, %long

          index = placeholder.field_value.length > 1 ? 1 : 0
          new_coord_info = get_new_coord_info_from_template(placeholder.field_value[index], params)
          placeholder.field_value[index] = new_coord_info
        end
      end

      def get_new_coord_info_from_template(value, template)
        value = value.to_s

        coord_info = {}
        coord_info[:lat]  =  {placeholder: "%lat", value: value.match(/lat=([\w\d\.\-]+),?/)}
        coord_info[:long] =  {placeholder: "%long", value: value.match(/long=([\w\d\.\-]+),?/)}
        coord_info[:alt]  =  {placeholder: "%alt", value: value.match(/alt=([\w\d\.\-]+),?/)}
        coord_info[:h_accuracy] = {placeholder: "%hAccuracy", value: value.match(/hAccuracy=([\w\d\.\-]+),?/)}
        coord_info[:v_accuracy] = {placeholder: "%vAccuracy", value: value.match(/vAccuracy=([\w\d\.\-]+),?/)}
        coord_info[:timestamp]  = {placeholder: "%timestamp", value: value.match(/timestamp=([\w\d\.\-:]+),?/)}

        result = template
        coord_info.each do |name, coord_replacement|
          replacement = coord_replacement[:value].present? ? coord_replacement[:value][1] : ""
          result = result.gsub(coord_replacement[:placeholder], replacement)
        end
        result
      end

    end
  end
end
