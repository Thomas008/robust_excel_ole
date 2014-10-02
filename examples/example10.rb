# example 10: save the sheets of a book as separate books

require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

require "fileutils"

include RobustExcelOle

ExcelApp.close_all
begin
  dir = '../spec/data/'
  file_name = dir + 'book_with_blank.xls'
  dir_save = 'C:/'
  file_name_save = dir_save + file_name
  book = Book.open(file_name)                   # open a book
  ExcelApp.reuse_if_possible.Visible = true     # make Excel visible 
  i = 0
  book.each do |sheet|
    i = i + 1
    file_name_save_sheet = file_name + "_sheet#{i}.xls"
    puts "file_name_save_sheet: #{file_name_save_sheet}"
	  # generate empty 
	  book_sheet = ExcelApp.Workbooks.Add
  end
 ensure                                                              
  ExcelApp.close_all                              # close workbooks, quit Excel application
 end

