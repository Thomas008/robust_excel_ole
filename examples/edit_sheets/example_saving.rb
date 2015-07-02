# example_saving.rb: 
# save the sheets of a book as separate workbooks

require 'rubygems'
require 'robust_excel_ole'
require "fileutils"

include RobustExcelOle

begin
  Excel.close_all
  dir = "C:/data"
  workbook_name = 'workbook.xls'
  base_name, suffix = workbook_name.split(".")
  file_name = dir + "/" + workbook_name
  excel = Excel.new(:visible => true)
  Book.unobtrusively(file_name, :if_closed => excel) do |book_orig| 
    book_orig.each do |sheet_orig|
      file_sheet_name = dir + "/" + base_name + "_" + sheet_orig.name + "." + suffix
      p "sheet.Name: #{sheet_orig.Name}"
      Excel.current.generate_workbook(file_sheet_name)
    end
  end
  Book.unobtrusively(file_name, :if_closed => excel) do |book_orig| 
    book_orig.each do |sheet_orig|
      p "sheet.Name: #{sheet_orig.Name}"
      file_sheet_name = dir + "/" + base_name + "_" + sheet_orig.name + "." + suffix
      # delete all existing sheets, and add the sheet    
      book = Book.open(file_sheet_name, :visible => true)
      book.add_sheet sheet_orig
      book.each do |sheet|
        sheet.Delete() unless sheet.Name == sheet_orig.Name 
      end
      book.close(:if_unsaved => :save)
      # alternative: delete all other sheets
      #book = Book.open(file_sheet_name, :force_excel => :new, :visible => true)
      #book.each do |sheet|
      #  p "sheet.Name: #{sheet.Name}"
      #  book[sheet.Name].Delete() unless sheet.Name == sheet_orig.Name
      #end
      #sleep 3
      #book.save_as(file_sheet_name, :if_exists => :overwrite)
      #book.close
    end
  end
end
