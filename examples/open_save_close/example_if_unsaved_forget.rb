# example_ifunsaved_forget.rb:
# open with :if_unsaved => :forget, :new_excel, close with :if_unsaved => :save 

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls' 
  book = Workbook.open(file_name)                      # open a book
  puts "file_name: #{file_name}"
  #book.excel.visible = true                        # make current Excel visible 
  sleep 1
  sheet = book.sheet(1)                                            # access a sheet
  first_cell = sheet[1,1].Value
  sheet[1,1] = first_cell == "simple" ? "complex" : "simple" # change a cell
  sleep 1
  puts "new_book: before"
  puts "file_name: #{file_name}"
  new_book = Workbook.open(file_name, :if_unsaved => :forget) # open another book with the same file name 
                                                          # and close the unsaved book without saving it
  puts "new_book: after"
  sheet_new_book = new_book.sheet(1)
  if (not book.alive?) && new_book.alive? && sheet_new_book[1,1].Value == first_cell then # check whether the unsaved book 
    puts "open with :if_unsaved => :forget : the unsaved book is closed and not saved."     # is closed and was not saved
  end
  sleep 1
  sheet_new_book[1,1] = sheet_new_book[1,1].Value == "simple" ? "complex" : "simple" # change a cell
  # open another book in a new Excel application, and make Excel visible, leaving the unsaved book open
  another_book = Workbook.open(file_name, :if_unsaved => :new_excel, :visible => true)  
  sleep 3                                                                  # leaving the unsaved book open  
  new_book.close(:if_unsaved => :forget )                                
  another_book.close
ensure
	  Excel.kill_all                                    # close all workbooks, quit Excel application
    rm_tmp(dir)
end