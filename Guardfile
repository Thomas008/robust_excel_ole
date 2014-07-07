# -*- mode: ruby -*-
# A sample Guardfile
# More info at https://github.com/guard/guard#readme

guard 'rspec', cli: "--color", all_after_pass: false, all_on_start: false do
  watch(%r{^spec/.+_spec\.rb$})
  watch(%r{^lib/robust_excel_ole/(.+)\.rb$})     { |m| "spec/#{m[1]}_spec.rb" }
  watch('lib/robust_excel_ole.rb')     { "spec" }
  watch('spec/spec_helper.rb')  { "spec" }
end
