module Word
  module ImageFunctions
    #
    #
    ## Image Functions
    #
    #

    def get_width_and_height_from_params
      edges = /(\d+)[xX](\d+)/.match(self.params)
      raise "Invalid params for image_size option on #{placeholder.field_identifier}. Expects [width]x[height], got #{self.params}." if edges.nil? || edges[1].nil? || edges[2].nil?
      width = edges[1].to_f
      height = edges[2].to_f
      [width,height]
    end

    def resize_image_answer(image, width, height)
      return image if image.nil? || (image.is_a?(String) && image == "") or width < 1 or height < 1
      return image if image.columns < width and image.rows < height
      image.resize([1.0 * width / image.columns, 1.0 * height / image.rows].min)
    end
  end
end
