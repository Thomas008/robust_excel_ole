# example_reuse.rb: open a book in a running Excel application and in a new one. make visible

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')

include RobustExcelOle

ExcelApp.close_all
begin
  dir = 'C:/'
  file_name1 = dir + 'simple.xls'
  file_name2 = dir + 'different_simple.xls'
  file_name3 = dir + 'different_simple.xls'
  file_name4 = dir + 'book_with_blank.xls'
  book1 = Book.open(file_name1)             # open a book in a new Excel application since no Excel is open
  ExcelApp.reuse.Visible = true # make Excel visible
  sleep 2
  book2 = Book.open(file_name2)             # open a new book in the same Excel application
  sleep 2                                   # (by default:  :reuse => true)
  book3 = Book.open(file_name3, :reuse => false, :visible => true) # open another book in a new Excel application, 
  sleep 2                                                          # make Excel visible
  book4 = Book.open(file_name4, :reuse => true, :visible => true)  # open anther book, and use a running Excel application
  sleep 2                                                          # (Excel chooses the first Excel application)        
  book1.close                               # close the books
  book2.close                                             
  book3.close
  book4.close                                         
ensure
  ExcelApp.close_all                       # close all workbooks, quit Excel application
end


