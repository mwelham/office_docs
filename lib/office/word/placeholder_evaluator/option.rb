module Word
  class Option
    attr_accessor :params, :placeholder

    def self.build_option_object(option_text, placeholder)
      option = option_text.split(':').map(&:strip).first.downcase
      params = option_text.split(':')[1..-1].join(':').strip

      #special case
      if /(\d+)x(\d+)/.match(option)
        params = option
        option = placeholder.is_map_answer? ? "map_size" : "image_size"
      end

      get_option_class(placeholder,option).new(placeholder, params)
    end

    def self.get_option_class(placeholder,option)
      begin
        option_class = "#{options_module}::#{option.camelize}".constantize
      rescue NameError
        raise "Unknown option #{option} used in the placeholder for #{placeholder.field_identifier}."
      end
    end

    def self.options_module
      "Word::Options"
    end

    def initialize(placeholder, params)
      self.placeholder = placeholder
      self.params = params
    end

    def apply_option
      raise 'implement me in the subclass'
    end

    protected

    #
    #
    ## Image Functions
    #
    #

    def get_width_and_height_from_params
      edges = /(\d+)x(\d+)/.match(self.params)
      raise "Invalid params for image_size option on #{placeholder.field_identifier}. Expects [width]x[height]."
      width = edges[1].to_f
      height = edges[2].to_f
      [width,height]
    end

    def resize_image_answer(image, width, height)
      return image if image.nil? or image == "" or width < 1 or height < 1
      return image if image.columns < width and image.rows < height
      image.resize([1.0 * width / image.columns, 1.0 * height / image.rows].min)
    end
  end
end
