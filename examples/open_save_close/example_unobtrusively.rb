# example_unobtrusively.rb:

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require File.join(File.dirname(__FILE__), '../../spec/helpers/create_temporary_dir')
require "fileutils"

include RobustExcelOle

Excel.close_all
begin
  dir = create_tmpdir
  simple_file = dir + 'workbook.xls'
  book = Book.open(simple_file, :visible => true)  # open a book, make Excel visible
  old_sheet = book[0]
  p "1st cell: #{old_sheet[1,1].value}"
  sleep 2
  Book.unobtrusively(simple_file) do |book|   # modify the book and keep its status unchanged
    sheet = book[0]
    sheet[1,1] = sheet[1,1].value == "simple" ? "complex" : "simple"
  end
  new_sheet = book[0]
  p "1st cell: #{new_sheet[1,1].value}"
  p "book saved" if book.Saved
  book.close                                 # close the book                      
ensure
  Excel.close_all                            # close all workbooks, quit Excel application
  rm_tmp(dir)
end
