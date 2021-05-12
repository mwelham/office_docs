Gem::Specification.new do |spec|
  spec.name        = "office_docs"
  spec.version     = "0.7.0"
  spec.summary     = "Manipulate Microsoft Office Open XML files"
  spec.description = "Generate and modify Word .docx and Excel .xlsx files"
  spec.authors     = ["Mike Welham", "Matthew Hirst"]
  spec.email       = "mikew@devicemagic.com"
  spec.files       = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^\.(git|ruby)|^(|test|spec|features)/}) }
  spec.homepage    = "https://github.com/devicemagic/office_docs"

  # Prevent pushing this gem to RubyGems.org. To allow pushes either set the 'allowed_push_host'
  # to allow pushing to a single host or delete this section to allow pushing to any host.
  if spec.respond_to?(:metadata)
    spec.metadata["allowed_push_host"] = ''
  else
    raise 'RubyGems 2.0 or newer is required to protect against public gem pushes.'
  end

  spec.add_dependency("nokogiri", ">= 1.10.10")
  spec.add_dependency("rmagick")
  spec.add_dependency("rubyzip", ">= 1.0.0")
  spec.add_dependency("activesupport")
  spec.add_dependency("actionview")
  spec.add_dependency("liquid", '3.0.6')
  spec.add_dependency("racc")
  spec.add_development_dependency("equivalent-xml", ">= 0.2.9")
  spec.add_development_dependency("pry")
  spec.add_development_dependency("rspec")
end
