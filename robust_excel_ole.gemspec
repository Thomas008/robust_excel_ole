# -*- coding: utf-8; mode: ruby -*-
$:.push File.expand_path("../lib", __FILE__)
require "robust_excel_ole/version"

Gem::Specification.new do |s|
  s.name        = "robust_excel_ole"
  s.version     = RobustExcelOle::VERSION
  s.authors     = ["traths"]
  s.email       = ["traths@cs.uni-potsdam.de"]
  s.homepage    = "https://github.com/Thomas008/robust_excel_ole"
  s.summary     = "RobustExcelOle is a wrapper library that specializes in the operation of Excel win32ole."
  s.description = "RobustExcelOle is to wrap the win32ole, and easy to use Excel operations with ruby. Detailed description please see the README."

  s.rubyforge_project = "robust_excel_ole"

  s.files         = `git ls-files`.split("\n")
  s.rdoc_options += [
                     '--main', 'README.rdoc',
                     '--charset', 'utf-8'
                    ]
  s.extra_rdoc_files = ['README.rdoc', 'LICENSE']

  s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
  s.add_development_dependency "rake", '>= 0.9.2'
  s.add_development_dependency "rspec", '>= 2.6.0'
  s.add_development_dependency "rb-fchange", '>= 0.0.5'
  s.add_development_dependency "wdm", '>= 0.0.3'
  s.add_development_dependency "win32console", '>= 1.3.2'
  s.add_development_dependency "guard-rspec", '>= 2.1.1'
  s.required_ruby_version = '>= 1.8.6'
end
