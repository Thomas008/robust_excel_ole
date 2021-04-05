# example_if_obstructed_close_if_saved.rb:
# open with :if_obstructed: :close_if_saved

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls'
  other_file_name = dir + 'more_data/workbook.xls'
  book = Workbook.open(file_name, :visible => true)  # open a workbook, make Excel visible
  sleep 1
  sheet = book.sheet(1)
  first_cell = sheet[1,1]                                   # access a worksheet
  sheet[1,1] = first_cell == "simple" ? "complex" : "simple"      # change a cell
  sleep 1
  begin
    new_book = Workbook.open(other_file_name, :if_obstructed => :close_if_saved) # raises an exception since the file is not saved
  rescue WorkbookBlocked => msg                                             
    puts "error: open: #{msg.message}"
  end                                                        
  book.save                                                           # save the unsaved workbook
  new_book = Workbook.open(file_name, :if_obstructed => :close_if_saved)  # open the new workbook, close the saved workbook    
  sleep 1
  new_sheet = new_book.sheet(1)
  new_first_cell = new_sheet[1,1].Value
  puts "the old book was saved" unless new_first_cell == first_cell 
  new_book.close                                 # close the workbook                  
ensure
  Excel.kill_all                         # close all workbooks, quit Excel application
  #rm_tmp(dir)
end
