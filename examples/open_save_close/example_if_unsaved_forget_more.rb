# example_ifunsaved_forget_more.rb:
# open with :if_unsaved => :forget, :new_excel, close with :if_unsaved => :save 

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls' 
  book = Workbook.open(file_name, :visible => true)  # open a workbook
  sheet = book.sheet(1)                                            # access a worksheet
  first_cell = sheet[1,1]
  sheet[1,1] = first_cell == "simple" ? "complex" : "simple" # change a cell
  sleep 1
  new_book = Workbook.open(file_name, :if_unsaved => :new_excel, :visible => true) # open another workbook with the same file name in a new Excel
  sheet_new_book = new_book.sheet(1)
  if (not book.alive?) && new_book.alive? && sheet_new_book[1,1] == first_cell then # check whether the unsaved workbook 
    puts "open with :if_unsaved => :forget : the unsaved book is closed and not saved."     # is closed and was not saved
  end
  sleep 1
  sheet_new_book[1,1] = sheet_new_book[1,1] == "simple" ? "complex" : "simple" # change a cell
  # open another workbook in the running Excel application, and make Excel visible, closing the unsaved workbook
  another_book = Workbook.open(file_name, :if_unsaved => :forget, :visible => true)  
  sleep 1
  sheet_another_book = another_book.sheet(1)
  sheet_another_book[1,1] = sheet_another_book[1,1] == "simple" ? "complex" : "simple" # change a cell                                                                   
  another_book.close(:if_unsaved => :forget )           # close the last workbook without saving it.                      
  book.close(:if_unsaved => :save)                      # close the first workbook and save it before
ensure
	  Excel.kill_all                                    # close all workbooks, quit Excel application
    rm_tmp(dir)
end