# introducing example example_open.rb

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

dir = create_tmpdir
simple_file = dir + 'workbook.xls'
different_simple_file = dir + 'different_workbook.xls'
another_simple_file = dir + 'another_workbook.xls'

Excel.kill_all

# open a workbook
book = Workbook.open(simple_file) 
puts "book: #{book}"
# open a workbook in a new Excel instance
book2 = Workbook.open(another_simple_file, :default => {:excel => :new})
puts "book2: #{book2}"
# open the workbook in a separate, reserved Excel instance.
book3 = Workbook.open(different_simple_file, :default => {:excel => :reserved_new})
puts "book3: #{book3}"
# open the workbook in a new Excel instance and make it visible
book4 = Workbook.open(simple_file, :force => {:excel => :new}, :visible => true)
puts "book4: #{book4}"
# open the workbook in a given Excel instance
excel1 = book.excel
book5 = Workbook.open(another_simple_file, :force => {:excel => excel1})
puts "book5: #{book5}"
# close a workbook
book.close
puts "close book -> book: #{book}"
# reopen the workbook
book6 = book.reopen
puts "reopened book: book6: #{book6}"
# unobtrusively opening a workbook
Workbook.unobtrusively(simple_file) do |book|
  sheet = book.sheet(1)
  sheet[1,1] = "c" 
end
book6.close
book5.close
book4.close
book3.close
book2.close
Excel.close_all
  