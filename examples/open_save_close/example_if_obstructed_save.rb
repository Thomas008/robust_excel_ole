# example_if_obstructed_save.rb:
# open with :if_obstructed: :save

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
  first_cell = sheet[1,1]                                        # access a worksheet
  sheet[1,1] = first_cell == "simple" ? "complex" : "simple"      # change a cell
  sleep 1
  new_book = Workbook.open(other_file_name, :if_obstructed => :save)  # open a workbook with the same file name in a different path
  sleep 1                                                         #save the old workbook, close it, before
  old_book = Workbook.open(file_name, :if_obstructed => :forget ,:visible => true) # open the old book    
  sleep 1
  old_sheet = old_book.sheet(1)
  old_first_cell = old_sheet[1,1]
  puts "the old book was saved" unless old_first_cell == first_cell 
  new_book.close                                 # close the workbooks                      
  old_book.close
ensure
  Excel.kill_all                         # close all workbooks, quit Excel application
  #rm_tmp(dir)
end
