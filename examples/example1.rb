# example 1: open a book, print the cells, rows, and columns of a sheet

require File.join(File.dirname(__FILE__), '../lib/robust_excel_ole')

module RobustExcelOle

	ExcelApp.close_all
	begin
	  filename = '../spec/data/simple.xls'
	  book = RobustExcelOle::Book.open(filename)
	  sheet = book[0]
	  cell = sheet[0,0]
	  i = 0
	  sheet.each do |cell|
	  	i = i + 1
	  	puts "#{i}. cell: #{cell.value}"
	  end
	  i = 0
	  sheet.each_row do |row|
	  	i = i + 1
	  	puts "#{i}. row: #{row.value}"
	  end
	  i = 0
	  sheet.each_column do |column|
	  i = i + 1
	  	puts "#{i}. column: #{column.value}"
	  end
	  sheet[0,0] = "complex"
	  book.save
	  book.close
	ensure
		  ExcelApp.close_all
	end

end
