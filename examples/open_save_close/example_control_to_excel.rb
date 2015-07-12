# example_give_control_to_excel.rb: 
# open, close, save  with giving control to Excel 

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls' 
  book = Book.open(file_name)          # open a book
  book.excel.visible = true                   # make current Excel visible 
  sleep 1
  sheet = book[0]                                                        # access a sheet
  sheet[1,1] = sheet[1,1].Value == "simple" ? "complex" : "simple"       # change a cell
  sleep 1
  begin
    new_book = Book.open(file_name, :if_unsaved => :alert) # open another book with the same file name 
  rescue ExcelUserCanceled => msg                          # if the user chooses not open the book,
  	puts "#{msg.message}"                                  #   an exeptions is raised
  end
  puts "new book has opened" if new_book
  Excel.current.Visible = true
  begin
  	book.close(:if_unsaved => :alert)                      # close the unsaved book. 
  rescue ExcelUserCanceled => msg                          # user is asked whether the unsaved book shall be saved
  	puts "#{msg.message}"                                  # if the user chooses to cancel, then an expeption is raised
  end
  if new_book then
  	begin
  	  new_book.save_as(file_name, :if_exists => :alert)    # save the new book, if it was opened
  	rescue ExcelErrorSave => msg                           # user is asked, whether the existing file shall be overwritten
  	  puts "save_as: #{msg.message}"                       # if the user chooses "no" or "cancel", an exception is raised
  	end 

  	new_book.close                                         # close the new book, if the user chose to open it
  end
ensure                                                              
  Excel.close_all                                       # close ALL workbooks, quit Excel application
  rm_tmp(dir)
end
