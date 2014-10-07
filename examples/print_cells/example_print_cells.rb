# example 1: open a book, print the cells, rows, and columns of a sheet

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')

require File.join(File.dirname(__FILE__), '../../spec/spec_helper')

require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  #dir = 'C:/'
  dir = create_tmpdir
  simple_file = dir + 'simple.xls'
  simple_save_file = dir + 'simple_save.xls'
  File.delete simple_save_file rescue nil
  book = Book.open(simple_file)
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
  Excel.close_all
end

  