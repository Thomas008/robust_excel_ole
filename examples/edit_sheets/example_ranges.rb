# example_ranges.rb: 
# access row and column ranges of a sheet. 

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all_known
begin
  dir = create_tmpdir
  simple_file = dir + 'workbook.xls'
  simple_save_file = dir + 'workbook_save.xls'
  File.delete simple_save_file rescue nil
  book = Book.open(simple_file)      # open a book
  sheet = book.sheet('Sheet1')             # access a sheet via the name
  row_r = sheet.row_range(1)         # access the whole range of the first row
  col_r = sheet.col_range(1, 1..2)   # access the first two cells of the range of the first column
  cell = col_r[0]                    # access the first cell of these cells 
  puts "row range of 1st row: #{row_r.values}"                     # puts the values of the first row
  puts "1st and 2nd cell of the 1st column : #{col_r.values}"      # and the first two cells of the first column
  puts "1st cell of these cells of the 1st columns: #{cell.Value}" # and the first cell of the row range of the 1st row 
  
  i = 0
  row_r.values.each do |value|          # access the values of the first row 
    i = i + 1
    puts "cell #{i} of the range of the 1st row: #{value}"
  end
 
  book.close
  
ensure
  Excel.close_all_known
  rm_tmp(dir)
end

  