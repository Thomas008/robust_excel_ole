# example_if_obstructed_close_if_saved.rb:
# open with :if_obstructed: :close_if_saved

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls'
  other_file_name = dir + 'more_data/workbook.xls'
  book = Workbook.open(file_name, :visible => true)  # open a book, make Excel visible
  sleep 1
  sheet = book.sheet(1)
  first_cell = sheet[1,1].Value                                   # access a sheet
  sheet[1,1] = first_cell == "simple" ? "complex" : "simple"      # change a cell
  sleep 1
  begin
    new_book = Workbook.open(other_file_name, :if_obstructed => :close_if_saved) # raises an exception since the file is not saved
    rescue WorkbookNotSaved => msg                                             
    puts "error: open: #{msg.message}"
  end                                                        
  book.save                                                           # save the unsaved book
  new_book = Workbook.open(file_name, :if_obstructed => :close_if_saved)  # open the new book, close the saved book    
  sleep 1
  new_sheet = new_book.sheet(1)
  new_first_cell = new_sheet[1,1].Value
  puts "the old book was saved" unless new_first_cell == first_cell 
  new_book.close                                 # close the books                  
ensure
  Excel.kill_all                         # close all workbooks, quit Excel application
  #rm_tmp(dir)
end
