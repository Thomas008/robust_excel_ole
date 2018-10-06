# introducing example example_range.rb

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

dir = create_tmpdir
simple_file = dir + 'another_workbook.xls'
simple_file2 = dir + 'workbook.xls'

Excel.kill_all

# opening a workbook
book = Workbook.open(simple_file, :visible => true)
# accessing the 1st worksheet of this workbook
sheet = book.sheet(1) 
# accessing a range consisting of one cell
range = sheet.range([1,2])
range2 = sheet.range("B1")
puts "range: #{range.Value}"
puts "range2: #{range2.Value}"
# the ranges are identical
# putting the values of the range
puts "range.values: #{range.values}"
# accessing a rectangular range given rows and columns
range3 = sheet.range([1..3,1..4])
range4 = sheet.range([1..3,"A".."D"])
# accessing the same range given the left top and the right bottum corner
range5 = sheet.range(["A1:D3"])
puts "range3: #{range3.Value}"
puts "range4: #{range4.Value}"
puts "range5: #{range5.Value}"
# the ranges are identival
# putting the values if the range
puts "range3.values: #{range3.values}"
# copying a range
range3.copy([6,2])
range6 = sheet.range([6..9,2..5])
puts "range3.Value: #{range3.Value}"
puts "range6.Value: #{range6.Value}"
# copying a range into another worksheet in another workbook of another Excel instance
book2 = Workbook.open(simple_file2, :excel => :new, :visible => true)
range3.copy([5,8], book2.sheet(3))
range7 = book2.sheet(3).range([5..8,8..11])
puts "range3.Value: #{range3.Value}"
puts "range7.Value: #{range7.Value}"
# adding a defined name referring to a range consisting of the first cell
book.add_name("name",[1,1])
# adding a defined name referring to a rectangular range
book.add_name("name",[1..2,3..4])
# assigning a value to that range 
book["name"] = [["foo", "bar"],[1.0, nil]]
# reading the value of that range
value = book["name"]
puts "value: #{value}" 
# renaming a range
book.rename_range("name", "new_name")
# deleting the name of a range
book.delete_name("new_name")
# reading the value of a cell
cell_value = sheet[1,1].Value
puts "cell_value: #{cell_value}"
# writing the value of a cell
sheet[1,1] = "bar"
Excel.close_all(:if_unsaved => :forget)
