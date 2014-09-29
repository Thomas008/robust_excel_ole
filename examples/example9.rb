# example 9: save the sheets of a book as separate books

require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

module RobustExcelOle

    ExcelApp.close_all
    begin
	ensure                                                              
  	  ExcelApp.close_all                                    # close workbooks, quit Excel application
	end

end
