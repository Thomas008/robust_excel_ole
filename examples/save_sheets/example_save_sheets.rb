# example_save_sheets.rb: 
# save the sheets of a book as separate books

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  suffix = '.xls'
  book_name = dir + 'book_with_blank' 
  book = Book.open(book_name + suffix)                   # open a book with several sheets
  book.visible true                                      # make Excel visible 
  i = 0
  book.each do |sheet|
    i = i + 1
    puts "#{i}. sheet:"
    sheet_name = book_name + "_sheet#{i}"
    puts "sheet_name: #{sheet_name}"
    Excel.generate_workbook absolute_path(sheet_name)   # generate an empty workbook
    sheet_book = Book.open(absolute_path(sheet_name) + suffix)  # open the book
    sheet_book.add_sheet sheet                                  # add the sheet
    sheet_book.save                                             # save it
    sheet_book.close                                            # close it
  end
  
 ensure                                                              
  Excel.close_all                              # close workbooks, quit Excel application
  rm_tmp(dir)
 end

