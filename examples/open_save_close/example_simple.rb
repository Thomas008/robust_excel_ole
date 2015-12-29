# example_simple.rb: 
# open a book, simple save, save_as, close

LOG_TO_STDOUT = false
REO_LOG_FILE = "reo2.log"
REO_LOG_DIR = ""

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  trace "hello"
  dir = create_tmpdir
  file_name = dir + 'workbook.xls'
  other_file_name = dir + 'different_workbook.xls'
  book = Book.open(file_name)                                # open a book.  default:  :read_only => false
  book.excel.visible = true                                  # make current Excel visible
  sheet = book[0]                                            # access a sheet
  sleep 1     
  sheet[1,1] = sheet[1,1].value == "simple" ? "complex" : "simple"  # change a cell
  sleep 1
  book.save                                                  # simple save
  begin
  	book.save_as(other_file_name)                            # save_as :  default :if_exists => :raise 
  rescue ExcelErrorSave => msg
  	puts "error: save_as: #{msg.message}"
  end
  book.save_as(other_file_name, :if_exists => :overwrite)    # save_as with :if_exists => :overwrite
  puts "save_as: saved successfully with option :if_exists => :overwrite"
  book.close                                                 # close the book
ensure
	  Excel.close_all                                         # close workbooks, quit Excel application
    rm_tmp(dir)
end		

