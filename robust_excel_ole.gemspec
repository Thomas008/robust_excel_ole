# -*- coding: utf-8; mode: ruby -*-
$:.push File.expand_path("../lib", __FILE__)
require "robust_excel_ole/version"

Gem::Specification.new do |s|
  s.name        = "robust_excel_ole"
  s.version     = RobustExcelOle::VERSION
  s.authors     = ["traths"]
  s.email       = ["Thomas.Raths@gmx.net"]
  s.homepage    = "https://github.com/Thomas008/robust_excel_ole"
  
  s.summary     = "RobustExcelOle automates processing Excel workbooks in Windows by using the win32ole library."
  s.description = "RobustExcelOle helps controlling Excel. 
                   It is designed to manage simultaneously running
                   Excel instances, even with simultanously happening user interactions.
                   This obviously includes standard tasks like reading and writing Excel workbooks.
                   RobustExcelOle deals with various cases of Excel (and user) behaviour,
                   supplies workarounds for some Excel bugs, and supports referenced libraries.
                   It runs on Windows and uses the win32ole library."                 

  s. licenses = ['MIT']
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
  s.add_development_dependency "rspec", '>= 2.6.0'
  s.required_ruby_version = '>= 1.8.6'
end
