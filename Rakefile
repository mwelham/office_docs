require 'rake/testtask'

Rake::TestTask.new do |t|
  t.libs << 'test'
end

desc "Run tests"
task :default => :test


Rake::TestTask.new("test:all") do |t|
  t.libs = ["lib", "test"]
  #t.warning = true
  t.test_files = FileList['test/**/test_*.rb']
end
