Gem::Specification.new do |spec|
  spec.name        = "office_docs"
  spec.version     = "0.3.4"
  spec.date        = "2013-04-26"
  spec.summary     = "Manipulate Microsoft Office Open XML files"
  spec.description = "Generate and modify Word .docx and Excel .xlsx files"
  spec.authors     = ["Mike Welham"]
  spec.email       = "mikew@devicemagic.com"
  spec.files       = ["lib/office_docs.rb"]
  spec.homepage    = "https://github.com/mwelham/office_docs"
  spec.add_dependency("nokogiri", ">= 1.5.2")
  spec.add_dependency("rmagick", ">= 2.12.2")
  spec.add_dependency("rubyzip", ">= 0.9.8")
  spec.add_development_dependency("equivalent-xml", ">= 0.2.9")
end
