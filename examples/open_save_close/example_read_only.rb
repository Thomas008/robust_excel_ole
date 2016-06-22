# example_read_only: open with read_only mode. save, close 

require File.expand_path('../../lib/robust_excel_ole', File.dirname(__FILE__))
require File.join(File.dirname(File.expand_path(__FILE__)), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls'
  other_file_name = dir + 'different_workbook.xls'
  book = Book.open(file_name, :read_only => true, :visible => true) # open a book with read_only and make Excel visible
  sheet = book.sheet(1)                                     			        # access a sheet
  sleep 1     
  sheet[1,1] = sheet[1,1].value == "simple" ? "complex" : "simple" # change a cell
  sleep 1
  begin
    book.save                                         # simple save. 
  rescue ExcelErrorSave => msg                        # raises an exception because book is opened in read_only mode
    puts "error: save_as: #{msg.message}"
  end
  book.close                                          # close the book without saving it 
ensure
  Excel.close_all                                     # close workbooks, quit Excel application
  rm_tmp(dir)
end

