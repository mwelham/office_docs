Gem::Specification.new do |spec|
  spec.name        = "office_docs"
  spec.version     = "0.3.9"
  spec.date        = "2014-04-24"
  spec.summary     = "Manipulate Microsoft Office Open XML files"
  spec.description = "Generate and modify Word .docx and Excel .xlsx files"
  spec.authors     = ["Mike Welham"]
  spec.email       = "mikew@devicemagic.com"
  spec.files       = ["lib/office_docs.rb"]
  spec.homepage    = "https://github.com/mwelham/office_docs"
  spec.add_dependency("nokogiri", ">= 1.5.2")
  spec.add_dependency("rmagick", ">= 2.12.2")
  spec.add_dependency("rubyzip", ">= 1.0.0")
  spec.add_development_dependency("equivalent-xml", ">= 0.2.9")
end
