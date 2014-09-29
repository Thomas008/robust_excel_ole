# example 8: open with :if_blocking_other:

require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

module RobustExcelOle

    ExcelApp.close_all
    begin
      dir = '../spec/data/'
	  file_name = dir + 'simple.xls'
	  other_dir = '../spec/data/more_data/'
	  other_file_name = other_dir + 'simple.xls'
	  book = RobustExcelOle::Book.open(file_name, :visible => true)     # open a book, make Excel application visible
	  sleep 3 
	  begin
	    new_book = RobustExcelOle::Book.open(other_file_name)   # open a book with the same file name in a different path
      rescue ExcelErrorOpen => msg                            # by default: raises an exception 
      	puts "open: #{msg.message}"
      end
      # open a new book with the same file name in a different path. close the old book before.
      new_book = RobustExcelOle::Book.open(other_file_name, :if_blocking_other => :forget) 
      # open another book with the same file name in a different path. Use a new Excel application
      sleep 3
      another_book = RobustExcelOle::Book.open(file_name, :if_blocking_other => :new_app, :visible => true)                                         
	  sleep 3
	  new_book.close                                        # close the books                      
	  another_book.close
	ensure
  	  ExcelApp.close_all                                    # close workbooks, quit Excel application
	end

end
