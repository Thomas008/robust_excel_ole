# example 1: open a book, print the cells, rows, and columns of a sheet

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
  book = Book.open(simple_file)
  sheet = book[0]
  cell = sheet[0,0]

  sheet_enum = proc do |enum_method|
    i = 0
    sheet.send(enum_method) do |cell|
      i = i + 1
      puts "sheet.#{enum_method} #{i}: #{cell.value}"
    end
  end

  sheet_enum[:each]
  sheet_enum[:each_row]
  sheet_enum[:each_column]

  col_r = sheet.col_range(0,1..2).values
  row_r = sheet.row_range(0,1..2).values
  puts "row range of 1st row, 1..2: #{row_r}"
  puts "column range of 1st column, 1..2: #{col_r}"
  sheet[0,0] = "complex"
  book.save
  book.close
ensure
  Excel.close_all
  rm_tmp(dir)
end

  