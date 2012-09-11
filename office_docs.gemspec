Gem::Specification.new do |spec|
  spec.name        = "office_docs"
  spec.version     = "0.3.0"
  spec.date        = "2012-07-20"
  spec.summary     = "Manipulate Microsoft Office Open XML files"
  spec.description = "Generate and modify Word .docx and Excel .xlsx files"
  spec.authors     = ["Mike Welham"]
  spec.email       = "mikew@devicemagic.com"
  spec.files       = ["lib/office_docs.rb"]
  spec.homepage    = "https://github.com/mwelham/office_docs"
  spec.add_dependency("nokogiri", ">= 1.5.2")
  spec.add_dependency("rmagick", ">= 2.12.2")
  spec.add_dependency("rubyzip", ">= 0.9.8")
end
