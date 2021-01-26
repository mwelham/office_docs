module Office
  class Error < StandardError; end

  class PackageError < Error; end
  class TypeError < Error; end
  class DocumentError < Error; end
  class LocatorError < Error; end
end
