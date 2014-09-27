# example 3: open a book, simple save, save_as, close

require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

module RobustExcelOle

    ExcelApp.close_all
    begin
   	  dir = '../spec/data/'
	  file_name = dir + 'simple.xls'
	  other_file_name = dir + 'different_simple.xls'
	  book = RobustExcelOle::Book.open(file_name)                # open a book.  default:  :read_only => false
	  ExcelApp.reuse_if_possible.Visible = true                  # make Excel visible
	  sheet = book[0]                                            # access a sheet
	  sleep 1     
	  sheet[0,0] = 
	    sheet[0,0].value == "simple" ? "complex" : "simple"        # change a cell
	  sleep 1
	  book.save                                                  # simple save
	  begin
	  	book.save_as(other_file_name)                            # save_as :  default :if_exists => :raise 
	  rescue ExcelErrorSave => msg
	  	puts "save_as error: #{msg.message}"
	  end
	  book.save_as(other_file_name, :if_exists => :overwrite)    # save_as with :if_exists => :overwrite
	  puts "save_as: saved successfully with option :if_exists => :overwrite"
	  book.close                                                 # close the book
	ensure
  	  ExcelApp.close_all                                         # close workbooks, quit Excel application
	end

end
