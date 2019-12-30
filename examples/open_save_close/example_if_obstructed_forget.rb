# example_if_obstructed_forget.rb: 
# open with :if_obstructed: :forget, :new_excel

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls'
  other_file_name = dir + 'more_data/workbook.xls'
  book = Workbook.open(file_name, :visible => true)  # open a workbook, make Excel application visible
  sleep 3 
  begin
    new_book = Workbook.open(other_file_name)        # open a workbook with the same file name in a different path
  rescue WorkbookBlocked => msg                   # by default: raises an exception 
  	puts "error: open: #{msg.message}"
  end
  # open a new book with the same file name in a different path. close the old book before.
  new_book = Workbook.open(other_file_name, :if_obstructed => :forget) 
  sleep 3
  # open another book with the same file name in a different path. Use a new Excel application
  another_book = Workbook.open(file_name, :if_obstructed => :new_excel, :visible => true)                                         
  sleep 3
  new_book.close                                 # close the workbooks                      
  another_book.close
ensure
  Excel.kill_all                         # close all workbooks, quit Excel application
  rm_tmp(dir)
end
