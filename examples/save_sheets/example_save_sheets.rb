# example 11: save the sheets of a book as separate books

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')

require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = 'C:/'
  file_name = dir + 'book_with_blank.xls'
  book = Book.open(file_name)                   # open a book
  Excel.current.Visible = true                  # make Excel visible 
  i = 0
  book.each do |sheet|
    i = i + 1
    puts "#{i}. sheet:"
    file_name_sheet = file_name + "_sheet#{i}.xls"
    puts "file_name_sheet: #{file_name_sheet}"
    empty_book = Excel.current.Workbooks.Add
    puts "class: #{empty_book}.class"
    sleep 2
    # empty_book.Workbooks(1).save_as(file_name_sheet) 
  end
  # 1. generate empty book
  #    book_sheet = Excel.Workbooks.Add
  # 2. copy with VBA method Copy (see Excel object model) a sheet into the empty book  
  
 ensure                                                              
  Excel.close_all                              # close workbooks, quit Excel application
 end

