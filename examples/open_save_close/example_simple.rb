# example_simple.rb: 
# open a book, simple save, save_as, close

LOG_TO_STDOUT = false
REO_LOG_FILE = "reo2.log"
REO_LOG_DIR = "C:/"

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls'
  other_file_name = dir + 'different_workbook.xls'
  book = Workbook.open(file_name, :visible => true)          # open a workbook.  default:  :read_only => false
  book.excel.visible = true                                  # make current Excel visible
  sheet = book.sheet(1)                                      # access a worksheet
  sleep 1     
  sheet[1,1] = sheet[1,1] == "simple" ? "complex" : "simple"  # change a cell
  sleep 1
  book.save                                                  # simple save
  begin
  	book.save_as(other_file_name)                            # save_as :  default :if_exists => :raise 
  rescue FileAlreadyExists => msg
  	puts "error: save_as: #{msg.message}"
  end
  book.save_as(other_file_name, :if_exists => :overwrite)    # save_as with :if_exists => :overwrite
  puts "save_as: saved successfully with option :if_exists => :overwrite"
  book.close                                                 # close the workbook
ensure
	  Excel.kill_all                                         # close workbooks, quit Excel application
    #rm_tmp(dir)
end		

