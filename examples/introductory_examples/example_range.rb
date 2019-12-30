# introducing example example_range.rb

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

dir = create_tmpdir
simple_file = dir + 'another_workbook.xls'
simple_file2 = dir + 'workbook.xls'

Excel.kill_all

# opening a workbook
puts "opening a workbook:"
book = Workbook.open(simple_file, :visible => true)
# accessing the 1st worksheet of this workbook
puts "accessing 1st worksheet"
sheet = book.sheet(1) 
# accessing a range consisting of one cell
range = sheet.range([1,2])
range2 = sheet.range("B1")
puts "cell [1,2]: #{range.Value}"
puts "B1: #{range2.Value}"
puts "the ranges are identical"
# the ranges are identical
# putting the values of the range
puts "range.values: #{range.values}"
# accessing a rectangular range given rows and columns
range3 = sheet.range([1..3,1..4])
range4 = sheet.range([1..3,"A".."D"])
# accessing the same range given the left top and the right bottum corner
range5 = sheet.range(["A1:D3"])
puts 'range([1..3,1..4]): #{range3.Value}'
puts 'range([1..3,"A".."D"]): #{range4.Value}'
puts 'range(["A1:D3"]): #{range5.Value}'
puts "the ranges are identical"
# the ranges are identical
# putting the values if the range
puts "range([1..3,1..4]).values: #{range3.values}"
# copying a range
puts "copying the range[1..3,1..4] to cell [6,2]"
range3.copy([6,2])
range6 = sheet.range([6..9,2..5])
puts "range([1..3,1..4]).Value: #{range3.Value}"
puts "range([6..9,2..5]).Value: #{range6.Value}"
# copying a range into another worksheet in another workbook of another Excel instance
puts "copying a range into another worksheet in another workbook of another Excel instance"
puts "opening a new workbook in a new Excel instance"
book2 = Workbook.open(simple_file2, :excel => :new, :visible => true)
range3.copy([5,8], book2.sheet(3))
range7 = book2.sheet(3).range([5..8,8..11])
puts "range([1..3,1..4]).Value: #{range3.Value}"
puts "new_book.sheet(3).range([5..8,8..11]).Value: #{range7.Value}"
# adding a defined name referring to a range consisting of the first cell
puts "adding a defined name 'name' referring to a range consisting of the 1st cell:"
book.add_name("name",[1,1])
# adding a defined name referring to a rectangular range
puts "adding a defined name 'name' referring to a rectangular range [1..2,3..4]:"
book.add_name("name",[1..2,3..4])
# assigning a value to that range 
puts "assigning a value to that range:"
book["name"] = [["foo", "bar"],[1.0, nil]]
# reading the value of that range
value = book["name"]
puts 'value of range "name": #{value}'
# renaming a range
puts 'renaming range from "name" to "new_name"'
book.rename_range("name", "new_name")
# deleting the name of a range
puts 'deleting name "new_name"'
book.delete_name("new_name")
# reading the value of a cell
cell_value = sheet[1,1].Value
puts "value of 1st cell: #{cell_value}"
# writing the value of a cell
puts 'writing the value "bar" into 1st cell'
sheet[1,1] = "bar"
Excel.kill_all
