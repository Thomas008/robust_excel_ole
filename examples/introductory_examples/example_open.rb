# introducing example example_open.rb

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

dir = create_tmpdir
simple_file = dir + 'workbook.xls'
different_simple_file = dir + 'different_workbook.xls'
another_simple_file = dir + 'another_workbook.xls'

Excel.kill_all

# open a workbook
puts "open a workbook"
book = Workbook.open(simple_file) 
puts "make it visible"
book.visible = true

ole_workbook = book.ole_workbook
puts "ole_workbook: #{ole_workbook}"
sheet = ole_workbook.sheet(1)
puts "sheet: #{sheet}"

using ToReoRefinement
puts "book: #{book}"
ole_workbook = book.ole_workbook
puts "ole_workbook: #{ole_workbook.inspect}"
upliftet_ole_workbook = ole_workbook.to_reo
puts "upliftet_ole_workbook: #{upliftet_ole_workbook.inspect}"
# open a workbook in a new Excel instance
puts "open the workbook in a new Excel instance and make it visible"
book2 = Workbook.open(another_simple_file, :default => {:excel => :new}, :visible => true)
puts "book2: #{book2}"
puts "create a new Excel"
excel1 = Excel.create(:visible => true)
# open the workbook in a given Excel instance
puts "open the workbook in a given Excel instance"
excel1 = book.excel
book3 = Workbook.open(another_simple_file, :force => {:excel => excel1})
puts "book3: #{book3}"
# close a workbook
puts "close the first workbook"
book.close
# reopen the workbook
puts "reopen the workbook"
book4 = book.open
puts "book4: #{book4}"
# unobtrusively opening a workbook
puts "unobtrusively opening the workbook"
Workbook.unobtrusively(simple_file) do |book|
  sheet = book.sheet(1)
  sheet[1,1] = "c" 
end
book4.close
book3.close
book2.close
Excel.kill_all
  