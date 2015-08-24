module Word
  class Option
    include Word::ImageFunctions

    attr_accessor :params, :placeholder

    SPECIAL_CASE_MAP = {
    }

    def self.build_option_object(option_text, placeholder)
      option = option_text.split(':').first.strip.downcase
      params = option_text.split(':')[1..-1].join(':').strip

      #special case
      if /(\d+)x(\d+)/.match(option)
        params = option
        option = placeholder.is_map_answer? ? "map_size" : "image_size"
      end

      get_option_class(placeholder,option).new(placeholder, params)
    end

    def self.get_option_class(placeholder,option)
      option_class = SPECIAL_CASE_MAP[option] || option.camelize

      begin
        option_class = "#{options_module}::#{option_class}".constantize
      rescue NameError
        #Ignoring bad options for now - maybe something to do in evaluate...
        #raise "Unknown option #{option} used in the placeholder for #{placeholder.field_identifier}"
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

  end
end
