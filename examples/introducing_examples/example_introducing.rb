# introducing example

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

dir = create_tmpdir
simple_file = dir + 'another_workbook.xls'

Excel.kill_all

# Let's open a workbook.
workbook = Workbook.open simple_file
# put name of the workbook
puts "name: #{workbook.Name}"
value = workbook['firstcell']
puts "value: #{value}"
# assigning a new value
workbook['firstcell'] = "new"
# saving the workbook
workbook.save
# closing the workbook
workbook.close
# reopening the workbook
workbook.reopen
# further operations
workbook['firstcell'] = "another_value"
# saved status of the workbook
puts "saved: #{workbook.Saved}"
# unobtrusively reading a workbook
Workbook.for_reading(simple_file) do |workbook|
  puts "value of first cell: #{workbook['firstcell']}"
end
puts "saved: #{workbook.Saved}"
# unobtrusively modifying a workbook
Workbook.for_modifying(simple_file) do |workbook|
  workbook['firstcell'] = "bar"
end
puts "saved: #{workbook.Saved}"
puts "value of first cell: #{workbook['firstcell']}"
