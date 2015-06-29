# example_saving.rb: 
# save the sheets of a book as separate workbooks

require 'rubygems'
require 'robust_excel_ole'
require "fileutils"

include RobustExcelOle

include RobustExcelOle

begin
  Excel.close_all
  dir = "C:/data"
  workbook_name = 'workbook.xls'
  base_name, suffix = workbook_name.split(".")
  file_name = dir + "/" + workbook_name
  excel = Excel.new(:reuse => false, :visible => true)
  Book.unobtrusively(file_name, :if_closed => excel, :keep_open => true) do |book_orig| 
    book_orig.each do |sheet_orig|
      book = Book.open(file_name, :force_excel => :new, :visible => true)
      book.each do |sheet|
        sheet.Delete() unless sheet.name == sheet_orig.name
      end
      file_sheet_name = dir + "/" + sheet_orig.name + "." + suffix
      book.save_as(file_sheet_name, :if_exists => :overwrite)
      book.close
      # does not work: error: cannot copy a sheet into another workbook 
      #Book.unobtrusively(file_sheet_name, :if_closed => :hidden) do |single_sheet_book| 
      #  single_sheet_book.add_sheet sheet
      #end
    end
  end
end
