require 'bundler/gem_tasks'
require 'rspec/core/rake_task'
require 'rake/testtask'

RSpec::Core::RakeTask.new

# prevent spec_helper from running the grammar build
ENV['NO_RAKE_GRAMMAR'] = 'true'

rule '.rb' => '.racc' do |t|
  sh "racc -o #{t.name} #{t.source}"
end

(lst = FileList['**/*.racc']).each do |raccfile|
  file raccfile.ext('.rb') => raccfile
end

desc 'build parser from grammar'
task grammar: lst.ext('.rb')

task :build => :grammar
task :release => :grammar
task :spec => :grammar
task :test => :grammar
task 'test:all' => :grammar

task :pry do
  require 'pry'
  $: << 'lib'
  require 'office_docs'
  Pry.start
end

Rake::TestTask.new do |t|
  t.libs << 'test'
end

desc 'Run specs and tests'
task default: %i[spec test]

# rake test TEST=test/test_foobar.rb TESTOPTS="--name=test_foobar1 -v"

Rake::TestTask.new("test:all") do |t|
  t.libs = ["lib", "test"]
  #t.warning = true
  t.test_files = FileList['test/**/test_*.rb']
end
