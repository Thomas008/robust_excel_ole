# example_save_sheets.rb: 
# save the sheets of a book as separate books

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')

require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = 'C:/'
  suffix = '.xls'
  book_name = dir + 'book_with_blank' 
  book = Book.open(book_name + suffix)                   # open a book
  Excel.current.Visible = true                           # make Excel visible 
  i = 0
  book.each do |sheet|
    i = i + 1
    puts "#{i}. sheet:"
    sheet_name = book_name + "_sheet#{i}"
    puts "sheet_name: #{sheet_name}"
    puts "absolute_path(sheet_name): #{absolute_path(sheet_name)}"
    # generate an empty workbook and save it
    excel = Excel.create
    excel.Workbooks.Add
    empty_workbook = excel.Workbooks.Item(1)
    empty_workbook.SaveAs(absolute_path(sheet_name), XlExcel8) 
    empty_workbook.Close
    # open the book, add the sheet and save it
    sheet_book = Book.open(absolute_path(sheet_name) + suffix)
    sheet_book.add_sheet sheet
    sheet_book.save
    sheet_book.close
  end
  
 ensure                                                              
  Excel.close_all                              # close workbooks, quit Excel application
 end

