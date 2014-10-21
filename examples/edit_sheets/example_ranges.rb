# example_ranges.rb: 
# access a sheet and ranges 

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  simple_file = dir + 'simple.xls'
  simple_save_file = dir + 'simple_save.xls'
  File.delete simple_save_file rescue nil
  book = Book.open(simple_file)      # open a book
  sheet = book['Sheet1']             # access a sheet via the name
  row_r = sheet.row_range(0)         # access the whole range of the first row
  col_r = sheet.col_range(0, 0..1)   # access the first two cells of the range of the first column
  puts "row range of 1st row: #{row_r.values}"                # puts the values of the first row
  puts "1st and 2nd cell of the 1st column : #{col_r.values}" # and the first two cells of the first column
  puts "1st cell of the 1st row: #{row_r[0].value}"           # and the second cell of the row range of the 1st row 
  
  i = 0
  row_r.values.each do |value|          # access the values of the first row
    i = i + 1
    puts "cell #{i} of the range of the 1st row: #{value}"
  end

  book.each do |sheet|               # access each sheet
    puts "sheet name: #{sheet.name}" #  put the sheet name
  end
  
  book.close
  
ensure
  Excel.close_all
  rm_tmp(dir)
end

  