# example_if_obstructed_forget.rb: 
# open with :if_obstructed: :forget, :new_app

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
  other_file_name = dir + 'more_data/simple.xls'
  book = Book.open(file_name, :visible => true)  # open a book, make Excel application visible
  sleep 3 
  begin
    new_book = Book.open(other_file_name)        # open a book with the same file name in a different path
  rescue ExcelErrorOpen => msg                   # by default: raises an exception 
  	puts "open: #{msg.message}"
  end
  # open a new book with the same file name in a different path. close the old book before.
  new_book = Book.open(other_file_name, :if_obstructed => :forget) 
  sleep 3
  # open another book with the same file name in a different path. Use a new Excel application
  another_book = Book.open(file_name, :if_obstructed => :new_app, :visible => true)                                         
  sleep 3
  new_book.close                                 # close the books                      
  another_book.close
ensure
  Excel.close_all                         # close all workbooks, quit Excel application
  rm_tmp(dir)
end
