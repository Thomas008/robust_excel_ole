# example 1: open a book, print the cells, rows, and columns of a sheet
require "fileutils"

require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

include RobustExcelOle

ExcelApp.close_all
begin
  simple_file = '../spec/data/simple.xls'
  simple_save_file = '../spec/data/simple_save.xls'
  File.delete @simple_save_file rescue nil
  FileUtils.copy simple_file, simple_save_file
  book = Book.open(simple_save_file)
  sheet = book[0]
  cell = sheet[0,0]
  i = 0
  sheet.each do |cell|
  	i = i + 1
  	puts "#{i}. cell: #{cell.value}"
  end
  i = 0
  sheet.each_row do |row|
  	i = i + 1
  	puts "#{i}. row: #{row.value}"
  end
  i = 0
  sheet.each_column do |column|
  i = i + 1
  	puts "#{i}. column: #{column.value}"
  end
  sheet[0,0] = "complex"
  book.save
  book.close
ensure
  ExcelApp.close_all
end

