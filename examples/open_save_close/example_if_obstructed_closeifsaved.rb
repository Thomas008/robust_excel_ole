# example_if_obstructed_close_if_saved.rb:
# open with :if_obstructed: :close_if_saved

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  file_name = dir + 'simple.xls'
  other_file_name = dir + 'more_data/simple.xls'
  book = Book.open(file_name, :visible => true)  # open a book, make Excel visible
  sleep 1
  sheet = book[0]
  first_cell = sheet[0,0].value                                   # access a sheet
  sheet[0,0] = first_cell == "simple" ? "complex" : "simple"      # change a cell
  sleep 1
  begin
    new_book = Book.open(other_file_name, :if_obstructed => :close_if_saved) # raises an exception since the file is not saved
    rescue ExcelErrorOpen => msg                                             
    puts "open: #{msg.message}"
  end                                                        
  book.save                                                           # save the unsaved book
  new_book = Book.open(file_name, :if_obstructed => :close_if_saved)  # open the new book, close the saved book    
  sleep 1
  new_sheet = new_book[0]
  new_first_cell = new_sheet[0,0].value
  puts "the old book was saved" unless new_first_cell == first_cell 
  new_book.close                                 # close the books                  
ensure
  Excel.close_all                         # close all workbooks, quit Excel application
  rm_tmp(dir)
end
