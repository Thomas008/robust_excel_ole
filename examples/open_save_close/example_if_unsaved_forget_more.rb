# example_ifunsaved_forget_more.rb:
# open with :if_unsaved => :forget, :new_app, close with :if_unsaved => :save 

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require "fileutils"
require 'tmpdir'

include RobustExcelOle

def create_tmpdir    
  tmpdir = Dir.mktmpdir
  FileUtils.cp_r(File.join(File.dirname(__FILE__), '../../spec/data'), tmpdir)
  tmpdir + '/data/'
end

def rm_tmp(tmpdir)    
  FileUtils.remove_entry_secure(File.dirname(tmpdir))
end

Excel.close_all
begin
  dir = create_tmpdir
  file_name = dir + 'simple.xls' 
  book = Book.open(file_name)                      # open a book
  Excel.current.Visible = true        # make Excel visible 
  sleep 1
  sheet = book[0]                                            # access a sheet
  first_cell = sheet[0,0].value
  sheet[0,0] = first_cell == "simple" ? "complex" : "simple" # change a cell
  sleep 1
  new_book = Book.open(file_name, :if_unsaved => :new_app, :visible => true) # open another book with the same file name in a new Excel
  sheet_new_book = new_book[0]
  if (not book.alive?) && new_book.alive? && sheet_new_book[0,0].value == first_cell then # check whether the unsaved book 
    puts "open with :if_unsaved => :forget : the unsaved book is closed and not saved."     # is closed and was not saved
  end
  sleep 1
  sheet_new_book[0,0] = sheet_new_book[0,0].value == "simple" ? "complex" : "simple" # change a cell
  # open another book in the running Excel application, and make Excel visible, closing the unsaved book
  another_book = Book.open(file_name, :if_unsaved => :forget, :visible => true)  
  sleep 1
  sheet_another_book = another_book[0]
  sheet_another_book[0,0] = sheet_another_book[0,0].value == "simple" ? "complex" : "simple" # change a cell                                                                   
  another_book.close(:if_unsaved => :forget )           # close the last book without saving it.                      
  book.close(:if_unsaved => :save)                      # close the first book and save it before
ensure
	  Excel.close_all                                    # close all workbooks, quit Excel application
    rm_tmp(dir)
end