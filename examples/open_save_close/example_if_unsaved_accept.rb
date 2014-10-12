# example_ifunsaved_accept.rb: 
# open with :if_unsaved => :accept, close with :if_unsaved => :save 

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  file_name = dir + 'simple.xls' 
  book = Book.open(file_name)                      # open a book 
  sheet = book[0]                                                  # access a sheet
  sheet[0,0] = sheet[0,0].value == "simple" ? "complex" : "simple" # change a cell
  begin
    new_book = Book.open(file_name)                # open another book with the same file name
  rescue ExcelErrorOpen => msg                     # by default: raises an exception:
    puts "open error: #{msg.message}"              # a book with the same name is already open and unsaved 
  end
  new_book = Book.open(file_name, :if_unsaved => :accept) # open another book with the same file name 
                                                          # and let the unsaved book open
  if book.alive? && new_book.alive? then                  # check whether the referenced workbooks
  	puts "open with :if_unsaved => :accept : the two books are alive." # respond to methods
  end
  if book == new_book then                                # check whether the book are equal
  	puts "both books are equal"
  end
  begin                                                                   
  	book.close                                          # close the book. by default: raises an exception:
  rescue ExcelErrorClose => msg                         #   book is unsaved
  	puts "close error: #{msg.message}"
  end
  book.close(:if_unsaved => :save)                      # save the book before closing it 
  puts "closed the book successfully with option :if_unsaved => :save"
  new_book.close                                        # close the other book. It is already saved.
ensure
	  Excel.close_all                                    # close workbooks, quit Excel application
    rm_tmp(dir)
end
