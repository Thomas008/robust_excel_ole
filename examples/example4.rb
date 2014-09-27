# example 4: open with read_only mode. save, close 

require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

module RobustExcelOle

    ExcelApp.close_all
    begin
   	  dir = '../spec/data/'
	  file_name = dir + 'simple.xls'
	  other_file_name = dir + 'different_simple.xls'
	  # open a book with read_only and make Excel visible
	  book = RobustExcelOle::Book.open(file_name, :read_only => true, :visible => true) 
	  sheet = book[0]                                     # access a sheet
	  sleep 1     
	  sheet[0,0] = 
	    sheet[0,0].value == "simple" ? "complex" : "simple" # change a cell
	  sleep 1
	  begin
	    book.save                                         # simple save. 
	  rescue ExcelErrorSave => msg                        # raises exception because book is opened in of read_only mode
	    puts "save_as error: #{msg.message}"
	  end
	  book.close                                          # close the book without saving it. 
	ensure
  	  ExcelApp.close_all                                  # close workbooks, quit Excel application
	end

end
