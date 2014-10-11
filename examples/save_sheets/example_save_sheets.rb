# example_save_sheets.rb: 
# save the sheets of a book as separate books

require File.join(File.dirname(__FILE__), '../../lib/robust_excel_ole')
require "fileutils"
require 'tmpdir'

include RobustExcelOle

def create_tmpdir    
  tmpdir = Dir.mktmpdir
  FileUtils.cp_r(File.join(File.dirname(__FILE__), '../../spec/data'), tmpdir)
  tmpdir + '/data/'
end

def rm_tmp(tmpdir)    
  FileUtils.remove_entry_secure(File.dirname(tmpdir))
end

Excel.close_all
begin
  dir = create_tmpdir
  suffix = '.xls'
  book_name = dir + 'book_with_blank' 
  book = Book.open(book_name + suffix)                   # open a book with several sheets
  Excel.current.Visible = true                           # make Excel visible 
  i = 0
  book.each do |sheet|
    i = i + 1
    puts "#{i}. sheet:"
    sheet_name = book_name + "_sheet#{i}"
    puts "sheet_name: #{sheet_name}"
    excel = Excel.create                   
    excel.Workbooks.Add                                          # generate an empty workbook
    empty_workbook = excel.Workbooks.Item(1)                     # save it
    empty_workbook.SaveAs(absolute_path(sheet_name), XlExcel8)   # close it
    empty_workbook.Close
    sheet_book = Book.open(absolute_path(sheet_name) + suffix)  # open the book
    sheet_book.add_sheet sheet                                  # add the sheet
    sheet_book.save                                             # save it
    sheet_book.close                                            # close it
  end
  
 ensure                                                              
  Excel.close_all                              # close workbooks, quit Excel application
  rm_tmp(dir)
 end

