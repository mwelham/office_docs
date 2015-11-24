module Word
  class Option
    include Word::ImageFunctions

    attr_accessor :params, :placeholder

    SPECIAL_CASE_MAP = {
    }

    # Use these to sort, so more important options are run first
    # Higher number is more important
    def importance
      10
    end

    def self.build_option_object(option_text, placeholder)
      option = option_text.split(':').first.strip.downcase
      params = option_text.split(':')[1..-1].join(':').strip

      #special case
      if /(\d+)x(\d+)/.match(option)
        params = option
        option = placeholder.is_map_answer? ? "map_size" : "image_size"
      end
      option_class = get_option_class(placeholder,option)
      if option_class.nil?
        nil
      else
        option_class.new(placeholder, params)
      end
    end

    def self.get_option_class(placeholder,option)
      option_class = SPECIAL_CASE_MAP[option] || option.camelize

      begin
        option_class = "#{options_module}::#{option_class}".constantize
      rescue NameError
        #raise "Unknown option #{option} used in the placeholder for #{placeholder.field_identifier}"

        #Ignoring bad options for now - maybe something to do in evaluate...
        return nil
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
