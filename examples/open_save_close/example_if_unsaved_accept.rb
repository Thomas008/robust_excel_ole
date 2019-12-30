# example_ifunsaved_accept.rb: 
# open with :if_unsaved => :accept, close with :if_unsaved => :save 

require_relative '../../lib/robust_excel_ole'
require_relative '../../spec/helpers/create_temporary_dir'
require "fileutils"

include RobustExcelOle

Excel.kill_all
begin
  dir = create_tmpdir
  file_name = dir + 'workbook.xls' 
  book = Workbook.open(file_name, :visible => true)                      # open a workbook 
  sheet = book.sheet(1)                                                  # access a worksheet
  sheet[1,1] = sheet[1,1].Value == "simple" ? "complex" : "simple" # change a cell
  begin
    new_book = Workbook.open(file_name)             # open another workbook with the same file name
  rescue WorkbookNotSaved => msg                    # by default: raises an exception:
    puts "error: open: #{msg.message}"              # a workbook with the same name is already open and unsaved 
  end
  new_book = Workbook.open(file_name, :if_unsaved => :accept) # open another workbook with the same file name 
                                                          # and let the unsaved workbook open
  if book.alive? && new_book.alive? then                  # check whether the referenced workbooks
  	puts "open with :if_unsaved => :accept : the two books are alive." # respond to methods
  end
  if book == new_book then                                # check whether the workbook are equal
  	puts "both books are equal"
  end
  begin                                                                   
  	book.close                                          # close the workbook. by default: raises an exception:
  rescue WorkbookNotSaved => msg                         #   workbook is unsaved
  	puts "close error: #{msg.message}"
  end
  book.close(:if_unsaved => :save)                      # save the book before closing it 
  puts "closed the book successfully with option :if_unsaved => :save"
  new_book.close                                        # close the other workbook. It is already saved.
ensure
	  Excel.kill_all                                    # close workbooks, quit Excel application
    rm_tmp(dir)
end
