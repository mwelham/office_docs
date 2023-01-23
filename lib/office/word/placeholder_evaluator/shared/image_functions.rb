module Word
  module ImageFunctions
    #
    #
    ## Image Functions
    #
    #

    def resample?
      !/noresample/.match(self.params)
    end

    def get_width_and_height_from_params
      edges = /(\d+)[xX](\d+)/.match(self.params)
      raise "Invalid params for image_size option on #{placeholder.field_identifier}. Expects [width]x[height], got #{self.params}." if edges.nil? || edges[1].nil? || edges[2].nil?
      width = edges[1].to_f
      height = edges[2].to_f
      [width,height]
    end

    def image_constraints(image, width, height)
      begin
        ratio = [1.0 * width / image.columns, 1.0 * height / image.rows].min
        {width: (image.columns * ratio).to_i, height: (image.rows * ratio).to_i}
      rescue
        {width: image.columns, height: image.rows}
      end
    end

    def resize_image_answer(image, width, height)
      return image if image.nil? || (image.is_a?(String) && image == "") || width < 1 || height < 1
      return image if image.columns < width && image.rows < height
      image.resize([1.0 * width / image.columns, 1.0 * height / image.rows].min)
    end
  end
end
